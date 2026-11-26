# -*- coding: utf-8 -*-
"""
ingesta.py — Consolidación sin cambios de lógica de:
  1. job_in_transit.py          -> main_in_transit()
  2. job_pedidos_transito.py    -> main_pedidos()
  3. backfill_historial_transito.py -> main_backfill()

 Parte 1 (in_transit):
  - Reporte Transito de la compania (Excel con conexion directa a su BD). El job lo refresca, lo guarda y recien
    despues lo lee. Reemplaza el scan de Hive.
  - Lake HU por lista literal IN de delivery order number (two-step, lookup por PK).
  - Hojas de salida en in_transit_peru.xlsx:
      bi          grain = linea del reporte (Key)   tal cual el reporte, nombres limpios, fechas tipadas
      dn_volumen  grain = delivery                  unidades del reporte + peso/volumen de Lake (el valor agregado);
                                                    no suma bultos de deliveries eliminadas (quedan en n_hu_eliminadas)
      hu_detail   grain = handling unit             HU tal cual Lake (todas sus columnas, valores y unidades, sin
                                                    convertir ni filtrar) + atributos de DOCUMENTO del reporte.
  - Historiales incrementales en parquet: historial_transito.parquet e historial_pedidos.parquet.

 Parte 2 (pedidos):
  - Parte del TRANSITO (hoja seguimiento: el bi + hito, llegada y salida a DC) y le pega, por SO + MTM, la 433 y
    el cliente de OrderMatch, buscando la SO en RPL_TRANSIT o RPL_SO (donde aparezca).
  - Grain = linea del bi (Key). Agrega la hoja pedidos_locales a in_transit_peru.xlsx.

 Parte 3 (backfill):
  - Lo que llego esta semana y ya salio del bi (aereos con aviso de llegada, maritimos o aereos con salida a DC):
    busca su HAWB/BL en los ultimos Inbound semanales de Elementos enviados y trae sus lineas y su 433 al historial.
"""

import logging
import re
import shutil
import tempfile
import time
import warnings
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import pyarrow.parquet as pq
import pythoncom                   # fail-fast: sin pywin32 revienta antes de tocar Excel o el lake
import win32com.client as win32

import lake_core as lc
import torre_config as C

# ══════════════════════════ CONFIG · IN-TRANSIT ══════════════════════════════
HOJA_REPORTE = None                # None = detecta la hoja cuyas columnas son todas Transito[...]; o el nombre exacto
TABLE_HU     = "lake.entregas_hu"
CHUNK_SIZE   = 1000

PREFIJO_COL = "Transito["        # todas las columnas del reporte vienen como Transito[Nombre]
COL_KEY  = C.COL_KEY
COL_DN   = C.COL_DN
COL_ITEM = C.COL_ITEM
COL_QTY  = C.COL_QTY
COLS_REQUERIDAS = [COL_KEY, COL_DN, COL_ITEM, COL_QTY, "so", "mot", "country",
                   "storage_loc", "bl_awb_num", "ship_date", "eta"]
COLS_RADAR = ["name_of_ocean_vessel", "equipment_num", "revised_eta", "outlook_eta"]   # paso 4 (A8): las pide el radar

# Fechas: el reporte mezcla serial de Excel (99999, 99999.999) con texto
COLS_FECHA_SERIAL = ["atd", "ata", "eta_warehouse", "ship_date", "pick_up",
                     "revised_eta", "outlook_eta", "etd"]
COLS_FECHA_TEXTO  = {"eta": "%m/%d/%Y",          # perfilado: patron 9/99/9999 => mes primero
                     "create_date": "%Y-%m-%d"}
MIN_COBERTURA_FECHA = 0.99         # bajo esto el formato cambio: revienta con el nombre de la columna

COBERTURA_HU_ALERTA = 0.50         # % de DN con HU en Lake bajo el cual se alerta (no revienta)

HU_NUMERIC_COLS = ["length", "width", "height", "net_weight", "total_weight",
                   "chargeable_weight", "volume_weight", "hu_quantity",
                   "total_volume", "tare_weight"]
HU_SUMABLES = ["hu_quantity", "net_weight", "total_weight", "chargeable_weight",
               "volume_weight", "total_volume"]

COLS_UNIDAD = ["unit_of_dimension", "volume_unit", "weight_unit", "weight_unit_tare"]

# Las 25 columnas de la tabla (perfiladas con diag_transito), listadas: si Lake quita una, la query
# revienta con su nombre; con SELECT * el archivo cambiaria de forma sin avisar
SQL_HU = f"""
SELECT delivery_number, handling_unit_id, sscc, batch_number,
       length, width, height, unit_of_dimension,
       total_volume, volume_unit,
       net_weight, total_weight, chargeable_weight, volume_weight, weight_unit,
       tare_weight, weight_unit_tare,
       quantity AS hu_quantity, delivery_deleted,
       ups_1z_number, return_1z_number_hu,
       gps_device_id, gps_device_vendor_id, gps_shipment_id,
       loc_updated_time
FROM {TABLE_HU}
WHERE delivery_number IN ({{in_list}})
"""

# Historial: lo que Transito borra queda aqui. Lo leen eventos y el correo para cruzar avisos de lo que ya salio
# Paso 5: el buque, el contenedor y la revised_eta tambien se guardan (sin ellos el radar pierde lo que ya arribo y
# Transito solto), y un cambio de ETA abre version nueva (asi queda la ETA anterior para la alerta de > 5 dias).
HIST_NUEVAS  = ["name_of_ocean_vessel", "equipment_num", "revised_eta"]   # una copia anterior al paso 5 no las trae
HIST_TEXTO   = [COL_KEY, "bl_awb_num", "so", "mtm", "mot", COL_DN, COL_ITEM, "country", "storage_loc",
                "name_of_ocean_vessel", "equipment_num"]
HIST_FECHAS  = ["ship_date", "outlook_eta", "revised_eta"]
HIST_VERSION = ["bl_awb_num", "so", "mtm", COL_QTY, "mot", "outlook_eta", "revised_eta"]   # si cambia uno: version nueva
HIST_COLS    = HIST_TEXTO + [COL_QTY] + HIST_FECHAS + ["visto_desde", "visto_hasta", "fuente"]
HIST_MAX_LOG = 15                  # HAWB/BL que salieron: cuantos se nombran en el log

# Pedidos: el OrderMatch suelta el 433 cuando la linea sale del transito (su match de transito se va con ella). Copia
# diaria de OM.xlsx: el correo saca de aqui el pedido de lo que ya salio. Grano = 422 x MTM x 433
PEDIDOS_LLAVE = ["so", "mtm", "so_local"]
PEDIDOS_COLS  = PEDIDOS_LLAVE + ["cliente", "sold_to", "visto_desde", "visto_hasta", "fuente"]   # sold_to: del 433

# ══════════════════════════ CONFIG · PEDIDOS ═════════════════════════════════
OM = "PEDIDOS_MATCH[{}]".format
C_SO_LOCAL, C_CLIENTE, C_MAT = OM("SO#"), OM("EndCustomerName"), OM("Material")
COLS_SO_TRANSITO = [OM("RPL_TRANSIT"), OM("RPL_SO")]      # donde puede estar la 422: se busca en las dos
RE_VARIAS_SO = r"[;,/|\s]+"

COLS_SEG = ["key", "so", "mtm", "transit_qty", "bl_awb_num", "mot", "hito", "llegada_fecha", "salida_dc_fecha",
            "llegada_aviso_mail", "salida_dc_mail", "corte_bi"]
COLS_VOLUMEN = ["bultos", "peso_kg", "volumen_m3", "volumen_exacto"]
COLUMNAS = ["so", "mtm", "transit_qty", *COLS_VOLUMEN, "bl_awb_num", "mot", "hito", "llegada_fecha", "salida_dc_fecha",
            "so_local", "cliente", "n_pedidos", "vinculo", "delivery_order_number", "segment", "key", "corte_bi"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
log_pedidos = logging.getLogger("pedidos_transito")

# ══════════════════════════ CONFIG · BACKFILL ════════════════════════════════
DIAS_AVISOS      = 7                        # avisos de llegada y salidas a DC de esta semana
ASUNTO_INBOUND   = "inbound & order match"  # se busca dentro del asunto, en minusculas
EXT_ADJUNTO      = (".xlsx", ".xlsm")       # todos los Excel del correo: el Inbound y el Order Match
MIN_DIGITOS      = 4                        # un HAWB/AWB/BL lleva numeros: 'LLEGADA' o 'CLIENTE' del asunto no se buscan
COLS_PEDIDO      = ["so", "mtm"]            # lo que cruza el 433 y el cliente: sin ellas la fila no sirve al correo
FILAS_ENCABEZADO = 10                       # los encabezados de una hoja pueden no estar en la fila 1
CLIENTE_OM       = ["endcustomername", "customer_name"]   # el cliente del Order Match del Inbound, en ese orden


# ══════════════════════════ HELPERS · IN-TRANSIT ═════════════════════════════
def limpiar_nombre(col: str) -> str:
    """'Transito[delivery order number]' -> 'delivery_order_number'."""
    interno = re.sub(r"^Transito\[(.*)\]$", r"\1", col)
    return re.sub(r"[^0-9a-z]+", "_", interno.lower()).strip("_")


def _cobertura(fecha: pd.Series, informada: pd.Series, nombre: str) -> None:
    cover = fecha[informada].notna().mean() if informada.any() else 1.0
    assert cover >= MIN_COBERTURA_FECHA, (
        f"{nombre}: solo {cover:.1%} parseable — cambio el formato en el reporte, re-perfilar")


def fecha_serial(serie: pd.Series, nombre: str) -> pd.Series:
    """Serial de Excel (sistema 1900). La parte decimal es la hora y se conserva."""
    numero = pd.to_numeric(serie, errors="coerce")
    fecha = pd.to_datetime(numero, unit="D", origin="1899-12-30")
    _cobertura(fecha, serie.notna(), nombre)
    return fecha


def fecha_texto(serie: pd.Series, formato: str, nombre: str) -> pd.Series:
    fecha = pd.to_datetime(serie.str.slice(0, 10), format=formato, errors="coerce")
    _cobertura(fecha, serie.notna(), nombre)
    return fecha


# ══════════════════════════ PARTE 1 · REPORTE TRANSITO ═════════════════════
def refrescar_reporte(ruta: Path) -> None:
    """Refresh sincronico en una instancia propia de Excel, guarda y cierra.
    Si algo falla, el job aborta aqui: nunca consulta Lake con las DN de la corrida anterior."""
    assert ruta.exists(), f"No existe el reporte {ruta} — revisar RUTA_REPORTE"
    mtime_antes = ruta.stat().st_mtime
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")        # instancia propia: no toca los Excel abiertos
    excel.Visible, excel.DisplayAlerts = False, False
    libro = None
    try:
        libro = excel.Workbooks.Open(str(ruta), UpdateLinks=0, ReadOnly=False)
        assert not libro.ReadOnly, f"{ruta.name} abrio en solo lectura: alguien lo tiene abierto — cerrarlo y re-correr"
        for conexion in libro.Connections:                # sin esto RefreshAll vuelve antes de terminar
            for tipo in ("OLEDBConnection", "ODBCConnection"):
                try:
                    getattr(conexion, tipo).BackgroundQuery = False
                except Exception:                         # cada conexion es de un solo tipo; el otro atributo no existe
                    pass
        t0 = time.time()
        libro.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        libro.Save()
        lc.log(f"    refresh OK en {time.time() - t0:.0f}s · {libro.Connections.Count} conexion(es)")
    finally:
        if libro is not None:
            libro.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
    assert ruta.stat().st_mtime > mtime_antes, (
        f"{ruta.name} no cambio tras el refresh — el dato seria el de la corrida anterior")


def hoja_del_reporte(ruta: Path) -> str:
    """El libro trae varias tablas (una por conexion). La hoja se elige por CONTENIDO, no por posicion:
    mover o agregar una pestana no puede hacer que el job lea otra tabla."""
    if HOJA_REPORTE is not None:
        return HOJA_REPORTE
    hojas = pd.ExcelFile(ruta).sheet_names
    candidatas = []
    for hoja in hojas:
        encabezado = pd.read_excel(ruta, sheet_name=hoja, nrows=0).columns
        if len(encabezado) and all(str(c).startswith(PREFIJO_COL) and str(c).endswith("]") for c in encabezado):
            candidatas.append(hoja)
    assert len(candidatas) == 1, (
        f"Se esperaba 1 hoja con columnas {PREFIJO_COL}...], hay {len(candidatas)}: {candidatas} "
        f"(hojas del libro: {hojas}) — fijar HOJA_REPORTE con el nombre exacto")
    lc.log(f"    hoja del reporte: {candidatas[0]!r} (de {len(hojas)} hojas)")
    return candidatas[0]


def cargar_reporte(ruta: Path) -> pd.DataFrame:
    crudo = pd.read_excel(ruta, sheet_name=hoja_del_reporte(ruta), dtype=str,
                          keep_default_na=False, na_values=[""])
    assert len(crudo), f"{ruta.name} sin filas tras el refresh — revisar la conexion del reporte"
    ajenas = [c for c in crudo.columns if not (str(c).startswith(PREFIJO_COL) and str(c).endswith("]"))]
    assert not ajenas, f"Columnas fuera del patron {PREFIJO_COL}...]: {ajenas} — cambio el reporte"

    rep = crudo.rename(columns=limpiar_nombre)
    assert rep.columns.is_unique, f"Nombres duplicados tras limpiar: {rep.columns[rep.columns.duplicated()].tolist()}"
    faltan = [c for c in COLS_REQUERIDAS if c not in rep.columns]
    assert not faltan, f"El reporte ya no trae {faltan} — renombraron columnas en la fuente"
    faltan_radar = [c for c in COLS_RADAR if c not in rep.columns]     # paso 4 (A8): se sabe aqui, no recien en el radar
    if faltan_radar:
        lc.log(f"    AVISO el reporte ya no trae {faltan_radar}: el radar maritimo saldra incompleto")
        C.AVISOS_PASO.append(f"Transito sin {faltan_radar} (radar)")

    # 38 columnas: apply por columna, no por fila
    rep = rep.apply(lambda s: s.astype("string").str.strip().where(~C.es_vacio_bi(s)))

    rep[COL_KEY] = C.norm_doc(rep[COL_KEY])
    rep[COL_DN] = C.norm_doc(rep[COL_DN])
    rep[COL_ITEM] = C.norm_doc(rep[COL_ITEM])
    dn_mal = rep[COL_DN].dropna().loc[lambda s: ~s.str.fullmatch(r"\d+")]
    assert dn_mal.empty, f"{len(dn_mal)} DN no numericas — Excel cambio el tipo de la columna"
    rep[COL_QTY] = pd.to_numeric(rep[COL_QTY], errors="raise")

    for col in COLS_FECHA_SERIAL:
        if col in rep.columns:
            rep[col] = fecha_serial(rep[col], col)
    for col, formato in COLS_FECHA_TEXTO.items():
        if col in rep.columns:
            rep[col] = fecha_texto(rep[col], formato, col)

    # grain: una fila = Key, y una fila = DN x item cuando hay DN
    assert rep[COL_KEY].notna().all() and rep[COL_KEY].is_unique, "Key nula o duplicada: el grain del reporte cambio"
    con_dn = rep[rep[COL_DN].notna()]
    dup = con_dn.duplicated([COL_DN, COL_ITEM], keep=False)
    assert not dup.any(), f"{dup.sum()} filas repiten DN x item — el merge con Lake multiplicaria HU"

    lc.log(f"    dominios · country={rep['country'].value_counts(dropna=False).to_dict()} "
           f"· mot={rep['mot'].value_counts(dropna=False).to_dict()}")
    desconocidos = sorted(set(rep["country"].dropna()) - {C.PAIS_PERU, C.PAIS_PANAMA})
    if desconocidos:                     # paso 5 (R9): antes un pais nuevo tumbaba toda la cadena de Peru; el filtro de
        lc.log(f"    AVISO country fuera de dominio {[C.PAIS_PERU, C.PAIS_PANAMA]}: {desconocidos} — quedan fuera "
               "del alcance (definir si entran)")  # alcance de abajo ya los deja fuera: se avisa y se sigue
        C.AVISOS_PASO.append(f"Transito trae country {desconocidos} (fuera de alcance)")

    es_panama = rep["country"].eq(C.PAIS_PANAMA)
    sloc = rep["storage_loc"].str.upper()
    lc.log(f"    Panama por storage_loc: {sloc[es_panama].value_counts(dropna=False).to_dict()}")
    en_alcance = rep["country"].eq(C.PAIS_PERU) | (es_panama & sloc.isin(C.SLOC_RAMA_PANAMA))
    if (~en_alcance).any():
        fuera = rep.loc[~en_alcance].groupby(["country", "storage_loc"], dropna=False).size().to_dict()
        lc.log(f"    fuera de alcance: {int((~en_alcance).sum())} lineas · country x storage_loc = {fuera}")
    rep = rep[en_alcance].copy()
    assert len(rep), "Sin lineas en alcance (PE + rama Panama) — revisar el reporte"
    return rep


# ══════════════════════════ PARTE 2 · LAKE ═══════════════════════════════════
def traer_hu(conn, keys: list) -> pd.DataFrame:
    hu = lc.query_chunked(conn, SQL_HU, keys, CHUNK_SIZE)
    assert not hu.empty, "Lake sin filas para ninguna DN — revisar formato de la llave (DN vs DC STO)"
    hu["delivery_number"] = C.norm_doc(hu["delivery_number"])
    dup = hu.duplicated(["delivery_number", "handling_unit_id"])
    assert not dup.any(), f"{dup.sum()} HU repetidas en Lake para la misma DN"

    for col in HU_NUMERIC_COLS:                           # Lake entrega todo varchar
        crudo = hu[col]
        hu[col] = pd.to_numeric(crudo.astype("string").str.strip(), errors="coerce")
        perdidos = int((hu[col].isna() & ~C.es_vacio_bi(crudo)).sum())
        if perdidos:
            lc.log(f"    AVISO {col}: {perdidos} valores no numericos quedaron vacios")

    unidades = {c: hu[c].astype("string").str.strip().str.upper().value_counts(dropna=False).to_dict()
                for c in COLS_UNIDAD}
    lc.log(f"    unidades Lake (sin convertir) · {unidades}")
    mezcladas = [c for c, dist in unidades.items() if len([u for u in dist if not pd.isna(u)]) > 1]
    if mezcladas:
        lc.log(f"    AVISO unidades mezcladas en {mezcladas}: las sumas de dn_volumen mezclan unidades "
               "(la vista embarques convierte a kg / m3)")
    return hu


def eliminadas(hu: pd.DataFrame) -> pd.Series:
    """Bultos de deliveries eliminadas en Lake. Siguen en hu_detail con su valor original; solo no se suman."""
    flag = hu["delivery_deleted"].astype("string").str.strip().str.upper()
    raros = sorted(set(flag.dropna()) - C.ELIMINADA_SI - C.ELIMINADA_NO - {""})
    assert not raros, f"delivery_deleted con valores no mapeados {raros} — definir si significan eliminada"
    return flag.isin(C.ELIMINADA_SI).fillna(False).astype(bool)


def atributos_documento(bi: pd.DataFrame):
    """Columnas del reporte constantes dentro de cada DN. Se detectan de la data: lo que varia
    por item (MTM, qty, Key) no viaja a hu_detail, asi cada HU aparece una sola vez."""
    con_dn = bi[bi[COL_DN].notna()]
    varia = con_dn.groupby(COL_DN).nunique(dropna=False).gt(1).any()
    cols_doc = [c for c in varia.index if not varia[c]]
    doc = con_dn.drop_duplicates(COL_DN)[[COL_DN] + cols_doc]
    return doc, sorted(varia[varia].index)


def construir_salidas(bi: pd.DataFrame, hu: pd.DataFrame):
    doc, cols_item = atributos_documento(bi)

    hu_detail = hu.merge(doc, left_on="delivery_number", right_on=COL_DN,
                         how="inner", validate="m:1").drop(columns=COL_DN)

    lineas = (bi[bi[COL_DN].notna()].groupby(COL_DN)
              .agg(lineas=(COL_KEY, "size"), unidades_reporte=(COL_QTY, "sum")))
    eliminada = eliminadas(hu)
    if eliminada.any():
        lc.log(f"    {int(eliminada.sum())} bultos de deliveries eliminadas: siguen en hu_detail, no suman en dn_volumen")
    vigentes = hu[~eliminada].groupby("delivery_number").agg(
        n_hu=("handling_unit_id", "nunique"), **{c: (c, "sum") for c in HU_SUMABLES})
    n_eliminadas = hu[eliminada].groupby("delivery_number").size().rename("n_hu_eliminadas")
    volumen = vigentes.join(n_eliminadas, how="outer")             # DN con solo eliminadas: n_hu vacio
    volumen["n_hu_eliminadas"] = volumen["n_hu_eliminadas"].fillna(0).astype(int)
    dn_volumen = (doc.merge(lineas, left_on=COL_DN, right_index=True, validate="1:1")
                     .merge(volumen, left_on=COL_DN, right_index=True, how="left", validate="1:1"))
    dn_volumen["tiene_hu"] = dn_volumen["n_hu"].notna()
    dn_volumen["dif_unidades_hu"] = dn_volumen["hu_quantity"] - dn_volumen["unidades_reporte"]  # control, no regla
    return dn_volumen, hu_detail, cols_item


# ══════════════════════════ HISTORIALES · PARQUET ════════════════════════════
def _tipar(df: pd.DataFrame) -> pd.DataFrame:
    """Mismos tipos en la foto y en el historial: la comparacion de versiones no ve '10' contra '10.0'."""
    df = df.copy()
    for c in HIST_TEXTO + ["fuente"]:
        if c in df.columns:
            df[c] = df[c].astype("string").str.strip()
    df[COL_QTY] = pd.to_numeric(df[COL_QTY], errors="coerce").astype("float64")
    for c in HIST_FECHAS + ["visto_desde", "visto_hasta"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def preparar_foto(rep: pd.DataFrame) -> pd.DataFrame:
    """Las columnas del historial a partir del bi, o de un Inbound semanal leido con cargar_reporte."""
    faltan = [c for c in [COL_KEY] + HIST_VERSION if c not in rep.columns]
    assert not faltan, f"El reporte no trae {faltan}: el historial no puede versionar la linea"
    foto = _tipar(rep.reindex(columns=HIST_TEXTO + [COL_QTY] + HIST_FECHAS)).reset_index(drop=True)
    assert foto[COL_KEY].notna().all() and foto[COL_KEY].is_unique, "Key nula o repetida en la foto"
    return foto


def historial_vacio() -> pd.DataFrame:
    vacio = preparar_foto(pd.DataFrame(columns=[COL_KEY] + HIST_VERSION))
    return _tipar(vacio.assign(visto_desde=pd.NaT, visto_hasta=pd.NaT, fuente=pd.NA))[HIST_COLS]


def leer_historial_transito() -> pd.DataFrame:
    """Irrecuperable: si existe y no se puede leer, se aborta. Nunca se sigue con vacio (lo reescribiria solo con hoy)."""
    if not C.HIST_PATH.exists():
        lc.log(f"    historial: {C.HIST_PATH.name} no existe, arranca con este corte")
        return historial_vacio()
    try:
        hist = pd.read_parquet(C.HIST_PATH)
    except Exception as e:
        raise RuntimeError(f"{C.HIST_PATH.name} existe pero no se pudo leer ({type(e).__name__}: {e}) — NO se reescribe: "
                           "restaurar una version desde el historial de versiones de OneDrive") from e
    faltan = [c for c in HIST_COLS if c not in hist.columns and c not in HIST_NUEVAS]
    assert not faltan, f"{C.HIST_PATH.name} perdio columnas {faltan} — alguien lo edito"
    nuevas = [c for c in HIST_NUEVAS if c not in hist.columns]
    if nuevas:                                          # copia anterior al paso 5: las columnas nuevas entran vacias
        lc.log(f"    historial: {C.HIST_PATH.name} sin {nuevas} (anterior al paso 5) — entran vacias desde hoy")
        hist = hist.assign(**{c: pd.NA for c in nuevas})
    return _tipar(hist)[HIST_COLS]


def aplicar_foto(hist: pd.DataFrame, rep: pd.DataFrame, corte: pd.Timestamp, fuente: str):
    """Suma una foto del in-transit al historial. Por linea de la foto: si esta igual que en el corte anterior se estira
    su visto_hasta; si es nueva, cambio, o reaparece despues de faltar, entra una fila nueva. Lo que estaba en el corte
    anterior y no viene en la foto salio (en aereo: llego al terminal). El mismo corte dos veces da lo mismo.
    Devuelve el historial y un resumen para el log."""
    foto = preparar_foto(rep)
    anterior = hist["visto_hasta"].max() if len(hist) else pd.NaT
    assert pd.isna(anterior) or corte >= anterior, (
        f"corte {corte:%d/%m/%Y %H:%M} anterior al ultimo del historial ({anterior:%d/%m/%Y %H:%M}): no se reescribe el pasado")
    hist = hist.reset_index(drop=True)
    ultima = hist.sort_values(["visto_desde", "visto_hasta"]).groupby(COL_KEY).tail(1)      # version vigente por linea
    previa = ultima.set_index(COL_KEY)[HIST_VERSION + ["visto_hasta"]].add_suffix("_prev")
    cruce = foto.join(previa, on=COL_KEY)
    conocida = cruce["visto_hasta_prev"].notna()
    igual = conocida.copy()
    for c in HIST_VERSION:
        texto = lambda s: s.astype("string").fillna("<vacio>")
        igual &= texto(cruce[c]).eq(texto(cruce[f"{c}_prev"]))
    sigue = igual & cruce["visto_hasta_prev"].ge(anterior)       # estaba en el corte anterior, con los mismos datos
    hist.loc[ultima.index[ultima[COL_KEY].isin(foto.loc[sigue, COL_KEY])], "visto_hasta"] = corte
    nuevas = _tipar(foto[~sigue].assign(visto_desde=corte, visto_hasta=corte, fuente=fuente))[HIST_COLS]
    if len(nuevas):
        hist = nuevas.reset_index(drop=True) if hist.empty else pd.concat([hist, nuevas], ignore_index=True)
    salieron = (ultima[ultima["visto_hasta"].ge(anterior) & ~ultima[COL_KEY].isin(foto[COL_KEY])]
                if pd.notna(anterior) else ultima.iloc[0:0])
    # Key estable: una DN x item que ya estaba con otra Key delata que el reporte renumero sus lineas
    dn_item = lambda d: d[COL_DN].fillna("") + "|" + d[COL_ITEM].fillna("")
    vistas = set(dn_item(ultima[ultima[COL_DN].notna()]))
    key_inestable = int(dn_item(foto[~conocida & foto[COL_DN].notna()]).isin(vistas).sum())
    return hist, {"lineas": len(foto), "nuevas": int((~conocida).sum()), "cambian": int((conocida & ~igual).sum()),
                  "reaparecen": int((igual & ~sigue).sum()), "salieron": salieron, "anterior": anterior,
                  "key_inestable": key_inestable}


def escribir_historial_transito(hist: pd.DataFrame) -> None:
    assert not hist.duplicated([COL_KEY, "visto_desde"]).any(), "historial con (Key, visto_desde) repetido"
    assert hist["visto_desde"].le(hist["visto_hasta"]).all(), "historial con visto_desde despues de visto_hasta"
    tmp = C.HIST_PATH.with_name(f"{C.HIST_PATH.stem}.tmp.parquet")
    hist[HIST_COLS].to_parquet(tmp, index=False)
    C.reemplazar_atomico(tmp, C.HIST_PATH)


def actualizar_historial(bi: pd.DataFrame, corte: pd.Timestamp) -> None:
    """El contraste con el corte anterior, guardado: lo que Transito borra queda en historial_transito.parquet."""
    hist, r = aplicar_foto(leer_historial_transito(), bi, corte, "bi")
    escribir_historial_transito(hist)
    previo = "primer corte" if pd.isna(r["anterior"]) else f"desde el corte del {r['anterior']:%d/%m %H:%M}"
    lc.log(f"    historial: {r['lineas']:,} lineas en el bi · {r['nuevas']} nuevas · {r['cambian']} cambiaron "
           f"(HAWB/BL, SO, MTM, cantidad o mot) · {r['reaparecen']} reaparecen · {len(r['salieron'])} salieron ({previo})")
    if len(r["salieron"]):
        salieron = r["salieron"]
        hawb = sorted(salieron["bl_awb_num"].dropna().unique())
        lc.log(f"      salieron · por mot {salieron['mot'].value_counts(dropna=False).to_dict()} · {len(hawb)} HAWB/BL: "
               f"{', '.join(hawb[:HIST_MAX_LOG])}{' …' if len(hawb) > HIST_MAX_LOG else ''}")
    if r["key_inestable"]:
        lc.log(f"    AVISO historial: {r['key_inestable']} Key nuevas traen una DN x item ya vista con otra Key — "
               "la Key no es estable entre refrescos: pasar el grano a DN x item")
    lc.log(f"    historial: {len(hist):,} filas · {hist[COL_KEY].nunique():,} lineas · desde {hist['visto_desde'].min():%d/%m/%Y}")


def leer_historial_pedidos() -> pd.DataFrame:
    """Igual que el de transito: si existe y no se puede leer, se aborta (no se reescribe solo con hoy)."""
    if not C.HIST_PEDIDOS_PATH.exists():
        return pd.DataFrame({c: pd.Series(dtype="datetime64[ns]" if c.startswith("visto") else "string")
                             for c in PEDIDOS_COLS})
    try:
        hist = pd.read_parquet(C.HIST_PEDIDOS_PATH)
    except Exception as e:
        raise RuntimeError(f"{C.HIST_PEDIDOS_PATH.name} existe pero no se pudo leer ({type(e).__name__}: {e}) — NO se "
                           "reescribe: restaurar una version desde el historial de versiones de OneDrive") from e
    if "sold_to" not in hist.columns:                  # copia de antes del 19/11: sold_to se llena en las siguientes
        hist = hist.assign(sold_to=pd.Series(pd.NA, index=hist.index, dtype="string"))
    faltan = [c for c in PEDIDOS_COLS if c not in hist.columns]
    assert not faltan, f"{C.HIST_PEDIDOS_PATH.name} perdio columnas {faltan} — alguien lo edito"
    return hist[PEDIDOS_COLS]


def columna_sold_to(columnas):
    """La columna del sold-to: 'PEDIDOS_MATCH[Sold_to]', 'Sold_to', 'Sold To'... (sin tabla, mayusculas ni simbolos)."""
    plano = lambda c: re.sub(r"[^0-9a-z]", "", re.sub(r"^.*\[(.*)\]$", r"\1", str(c)).lower())
    return next((c for c in columnas if plano(c) == "soldto"), None)


def guardar_pedidos(pedidos: pd.DataFrame, corte: pd.Timestamp, fuente: str) -> int:
    """Suma pedidos (so 422, mtm, so_local 433, cliente y, si viene, sold_to) a historial_pedidos.parquet. Por 422 x MTM x
    433 queda el primer y el ultimo corte en que se vio y el ultimo cliente y sold-to conocidos: correrlo dos veces da
    lo mismo. Devuelve cuantos son nuevos."""
    faltan = [c for c in PEDIDOS_LLAVE + ["cliente"] if c not in pedidos.columns]
    assert not faltan, f"pedidos sin {faltan}"
    previo = leer_historial_pedidos()
    vistos = (pedidos.reindex(columns=PEDIDOS_LLAVE + ["cliente", "sold_to"]).dropna(subset=PEDIDOS_LLAVE)
              .assign(visto_desde=corte, visto_hasta=corte, fuente=fuente))
    todo = vistos if previo.empty else pd.concat([previo, vistos], ignore_index=True)
    for c in PEDIDOS_LLAVE + ["cliente", "sold_to", "fuente"]:
        todo[c] = todo[c].astype("string").str.strip()
    hist = (todo.assign(visto_desde=pd.to_datetime(todo["visto_desde"]), visto_hasta=pd.to_datetime(todo["visto_hasta"]))
                .sort_values("visto_hasta", kind="stable")
                .groupby(PEDIDOS_LLAVE, as_index=False)
                .agg(cliente=("cliente", "last"), sold_to=("sold_to", "last"), visto_desde=("visto_desde", "min"),
                     visto_hasta=("visto_hasta", "max"), fuente=("fuente", "last")))
    assert not hist.duplicated(PEDIDOS_LLAVE).any(), "historial de pedidos con 422 x MTM x 433 repetido"
    tmp = C.HIST_PEDIDOS_PATH.with_name(f"{C.HIST_PEDIDOS_PATH.stem}.tmp.parquet")
    hist[PEDIDOS_COLS].to_parquet(tmp, index=False)
    C.reemplazar_atomico(tmp, C.HIST_PEDIDOS_PATH)
    return len(hist) - len(previo)


def copiar_pedidos(corte: pd.Timestamp) -> None:
    """Copia del OrderMatch de hoy (misma lectura y cruce que el correo). Si falla, el bi y el historial de transito ya
    quedaron escritos y el correo usa la ultima copia: se avisa y se sigue."""
    try:
        om = leer_ordermatch()
        pedidos = pedidos_de_om(om)
        col = columna_sold_to(om.columns)
        if col:                                        # el sold-to es del pedido: uno por 433
            sold = om[[C_SO_LOCAL, col]].set_axis(["so_local", "sold_to"], axis=1)
            sold = sold.assign(so_local=C.txt(sold["so_local"]), sold_to=C.txt(sold["sold_to"])).dropna()
            pedidos = pedidos.merge(sold.drop_duplicates("so_local"), on="so_local", how="left")
        else:
            lc.log(f"    AVISO el OrderMatch no trae sold-to (columnas con 'sold': "
                   f"{[c for c in om.columns if 'sold' in str(c).lower()]}) — se guarda sin el")
        nuevos = guardar_pedidos(pedidos, corte, "om")
        lc.log(f"    historial de pedidos: {len(pedidos):,} 422 x MTM -> 433 en el OrderMatch de hoy · {nuevos} nuevos "
               f"· sold-to {'de ' + str(col) if col else 'no'} · {C.HIST_PEDIDOS_PATH.name}")
    except Exception as e:
        lc.log(f"    ERROR {C.HIST_PEDIDOS_PATH.name} sin la copia de hoy ({type(e).__name__}: {e}) — "
               "el correo usa la ultima copia para lo que salio del transito")
        C.AVISOS_PASO.append(f"{C.HIST_PEDIDOS_PATH.name} sin la copia de hoy ({type(e).__name__})")   # paso 4 (A1)


def main_in_transit():
    lc.log("IN-TRANSIT v2 — parte 1: refresh del reporte Transito")
    refrescar_reporte(C.RUTA_REPORTE)
    corte = pd.Timestamp.now().floor("s")                     # el reporte refleja este momento: corte del historial
    bi = cargar_reporte(C.RUTA_REPORTE)

    sin_dn = int(bi[COL_DN].isna().sum())
    keys = sorted(bi[COL_DN].dropna().unique())
    assert keys, "Ninguna linea con delivery order number — nada que cruzar contra Lake"
    lc.log(f"    bi: {len(bi):,} lineas · {len(keys):,} DN · {sin_dn} lineas sin DN (quedan en bi, no van a Lake)")

    if C.USAR_LAKE:
        lc.log(f"IN-TRANSIT v2 — parte 2: {len(keys):,} DN -> Lake")
        conn = lc.get_connection()
        hu = traer_hu(conn, keys)

        dn_volumen, hu_detail, cols_item = construir_salidas(bi, hu)
        cobertura = dn_volumen["tiene_hu"].mean()
        lc.log(f"    dn_volumen: {len(dn_volumen):,} DN · con HU {cobertura:.0%} · "
               f"hu_detail: {len(hu_detail):,} HU")
        lc.log(f"    columnas de item (no viajan a hu_detail): {cols_item}")
        if cobertura < COBERTURA_HU_ALERTA:
            lc.log(f"    AVISO cobertura HU {cobertura:.0%} < {COBERTURA_HU_ALERTA:.0%} — "
                   f"probar dc_sto como llave antes de asumir que faltan HU")
        descuadre = int((dn_volumen["tiene_hu"] & dn_volumen["dif_unidades_hu"].ne(0)).sum())
        lc.log(f"    control: {descuadre} DN con unidades HU distintas a las del reporte")
        hojas_salida = {"bi": bi, "dn_volumen": dn_volumen, "hu_detail": hu_detail}
    else:
        hojas_salida = {"bi": bi}

    lc.atomic_to_excel(
        hojas_salida, C.OUT_PATH,
        f"reporte Transito (refresh {pd.Timestamp.now():%Y-%m-%d %H:%M}) "
        f"· alcance={C.PAIS_PERU} + {C.PAIS_PANAMA} en {C.SLOC_RAMA_PANAMA} · two-step Lake por {COL_DN}", None)
    actualizar_historial(bi, corte)                           # despues del bi: si falla, el bi de hoy ya quedo escrito
    copiar_pedidos(corte)
    lc.log("IN-TRANSIT v2 DONE")


# ══════════════════════════ LECTURA Y CRUCE · PEDIDOS ════════════════════════
def leer_seguimiento():
    assert C.RUTA_IN_TRANSIT.exists(), f"No existe {C.RUTA_IN_TRANSIT} — correr job_in_transit y job_eventos_transito"
    mtime = C.RUTA_IN_TRANSIT.stat().st_mtime
    with pd.ExcelFile(C.RUTA_IN_TRANSIT) as libro:
        if C.HOJA_SEGUIMIENTO not in libro.sheet_names:
            raise RuntimeError(f"{C.RUTA_IN_TRANSIT.name} sin hoja {C.HOJA_SEGUIMIENTO!r} — correr job_eventos_transito")
        seg = pd.read_excel(libro, C.HOJA_SEGUIMIENTO,
                            dtype={"key": str, "so": str, "mtm": str, "bl_awb_num": str, "delivery_order_number": str})
    faltan = [c for c in COLS_SEG if c not in seg.columns]
    assert not faltan, f"{C.HOJA_SEGUIMIENTO} ya no trae {faltan} — cambio job_eventos_transito"
    corte = pd.Timestamp(seg["corte_bi"].iloc[0])
    if C.edad_horas(corte) > C.MAX_HORAS:
        raise RuntimeError(f"seguimiento con corte {corte:%d/%m %H:%M} — correr job_in_transit y job_eventos_transito")
    return seg, mtime, corte


def leer_ordermatch() -> pd.DataFrame:
    assert C.RUTA_OM.exists(), f"No existe {C.RUTA_OM}"
    edad = C.edad_horas(datetime.fromtimestamp(C.RUTA_OM.stat().st_mtime))
    if edad > C.MAX_HORAS:
        raise RuntimeError(f"{C.RUTA_OM.name} tiene {edad:.0f} h sin refresh — correr job_in_transit (lo refresca)")
    with pd.ExcelFile(C.RUTA_OM) as libro:
        hojas = [h for h in libro.sheet_names if C_SO_LOCAL in pd.read_excel(libro, h, nrows=0).columns]
        assert len(hojas) == 1, f"Se esperaba 1 hoja con {C_SO_LOCAL!r} en {C.RUTA_OM.name}, hay {hojas}"
        om = pd.read_excel(libro, hojas[0], dtype=str, keep_default_na=False, na_values=[""])
    faltan = [c for c in [C_SO_LOCAL, C_CLIENTE, C_MAT] + COLS_SO_TRANSITO if c not in om.columns]
    assert not faltan, f"OrderMatch ya no trae {faltan} — renombraron columnas en el modelo"
    return om


def leer_hu_detail():
    """Bultos de Lake tal cual (job_in_transit). Sin la hoja, las lineas quedan sin volumen y se avisa."""
    with pd.ExcelFile(C.RUTA_IN_TRANSIT) as libro:
        if C.HOJA_HU not in libro.sheet_names:
            log_pedidos.warning("%s sin hoja %r: las lineas quedan sin volumen", C.RUTA_IN_TRANSIT.name, C.HOJA_HU)
            return None
        return pd.read_excel(libro, C.HOJA_HU, dtype={"delivery_number": str, "handling_unit_id": str,
                                                     "delivery_deleted": str})


def volumen_por_linea(seg: pd.DataFrame, hu) -> pd.DataFrame:
    """Bultos, kg y m3 de cada linea del bi (Key) desde los bultos de su DN.
    DN con una sola linea -> exacto. DN con varias lineas (varios items / MTM en la misma entrega): Lake da el
    volumen por DN, no por MTM -> se reparte por unidades de cada linea (volumen_exacto = False).
    Se reparte entre TODAS las lineas de la DN (tambien las de stock), asi la suma siempre da la DN."""
    vacio = pd.DataFrame(index=pd.Index(seg["key"], name="key"), columns=COLS_VOLUMEN + ["dn", "n_lineas_dn"])
    if hu is None or hu.empty:
        return vacio
    v = C.en_unidades_estandar(hu)                             # solo bultos vigentes, con su unidad convertida
    suma = lambda x: x.sum(min_count=1)                        # DN sin ninguna medida: vacio, no 0
    por_dn = (v.assign(dn=C.txt(v["delivery_number"])).groupby("dn")
               .agg(bultos=("dn", "size"), peso_kg=("peso_bruto_kg", suma), volumen_m3=("volumen_m3", suma)))
    lineas = pd.DataFrame({"key": seg["key"], "dn": C.txt(seg["delivery_order_number"]),
                           "qty": pd.to_numeric(seg["transit_qty"], errors="coerce")})
    g = lineas.groupby("dn")
    n_lineas, qty_dn = g["key"].transform("size"), g["qty"].transform("sum")
    parte = (lineas["qty"] / qty_dn).where(qty_dn > 0, 1 / n_lineas)   # sin unidades: partes iguales
    lineas = lineas.join(por_dn, on="dn")
    for col in ("bultos", "peso_kg", "volumen_m3"):
        lineas[col] = lineas[col] * parte
    lineas["volumen_exacto"] = n_lineas.eq(1) & lineas["bultos"].notna()
    lineas["n_lineas_dn"] = n_lineas                          # para quien agrupe lineas: si junta toda la DN, es exacto
    return lineas.set_index("key")[COLS_VOLUMEN + ["dn", "n_lineas_dn"]]


def pedidos_de_om(om: pd.DataFrame) -> pd.DataFrame:
    """SO de transito x MTM -> 433 + cliente. Una fila por combinacion: los items del mismo pedido se juntan."""
    partes = [pd.DataFrame({"so": C.txt(om[col]).str.split(RE_VARIAS_SO, regex=True),   # celda con varias SO: todas
                            "mtm": C.txt(om[C_MAT]).str.upper(), "so_local": C.txt(om[C_SO_LOCAL]),
                            "cliente": om[C_CLIENTE].astype("string").str.strip()}).explode("so")
              for col in COLS_SO_TRANSITO]
    largo = pd.concat(partes, ignore_index=True)
    largo["so"] = C.txt(largo["so"])
    return largo.dropna(subset=["so", "so_local"]).drop_duplicates(["so", "mtm", "so_local"])


def lineas_con_pedido(seg: pd.DataFrame, pedidos: pd.DataFrame) -> pd.DataFrame:
    """Lineas del transito x pedido: una linea sin pedido queda una vez, con so_local vacio."""
    s = seg.assign(so=C.txt(seg["so"]), mtm=C.txt(seg["mtm"]).str.upper(), bl_awb_num=C.txt(seg["bl_awb_num"]),
                   transit_qty=pd.to_numeric(seg["transit_qty"], errors="coerce"))
    largo = s.merge(pedidos, on=["so", "mtm"], how="left")
    so_en_om = largo["so"].isin(set(pedidos["so"]))
    largo["vinculo"] = np.select([largo["so_local"].notna(), so_en_om],
                                 ["con pedido", "SO en OrderMatch con otro MTM"], default="sin pedido (stock)")
    return largo


def por_linea(largo: pd.DataFrame) -> pd.DataFrame:
    """Vuelve al grain del bi: una fila por Key."""
    unir = lambda x: " | ".join(dict.fromkeys(x.dropna().astype(str)))      # sin repetir, en orden
    g = largo.groupby("key", sort=False)
    pedidos = pd.DataFrame({"so_local": g["so_local"].agg(unir), "cliente": g["cliente"].agg(unir),
                            "n_pedidos": g["so_local"].nunique()})
    base = largo.drop_duplicates("key").set_index("key").drop(columns=["so_local", "cliente"])
    vista = base.join(pedidos).reset_index()
    vista[["so_local", "cliente"]] = vista[["so_local", "cliente"]].replace("", pd.NA)
    return vista.reindex(columns=[c for c in COLUMNAS if c in vista.columns or c == "key"])


def escribir(vista: pd.DataFrame, mtime_leido: float) -> None:
    C.reemplazar_hojas_excel(C.RUTA_IN_TRANSIT, {C.HOJA_PEDIDOS: vista}, mtime_leido, C.HOJAS_AJENAS_PEDIDOS)


def main_pedidos() -> None:
    log_pedidos.info("=== PEDIDOS · linea en transito -> SO + MTM -> 433 + cliente ===")
    seg, mtime, _ = leer_seguimiento()
    vista = por_linea(lineas_con_pedido(seg, pedidos_de_om(leer_ordermatch())))
    assert len(vista) == len(seg), f"la vista tiene {len(vista)} filas y el transito {len(seg)}: el cruce multiplico"
    volumen = volumen_por_linea(seg, leer_hu_detail())
    vista = vista.drop(columns=[c for c in COLS_VOLUMEN if c in vista]).join(volumen[COLS_VOLUMEN], on="key")
    vista = vista.reindex(columns=[c for c in COLUMNAS if c in vista.columns])
    vista = vista.round({"bultos": 2, "peso_kg": 1, "volumen_m3": 4})        # repartidos: sin colas de decimales
    log_pedidos.info("volumen: %d lineas exactas · %d repartidas por unidades (DN con varias lineas) · %d sin bultos",
                     int(vista["volumen_exacto"].eq(True).sum()),
                     int((vista["volumen_exacto"].eq(False) & vista["bultos"].notna()).sum()),
                     int(vista["bultos"].isna().sum()))
    escribir(vista, mtime)
    log_pedidos.info("pedidos_locales: %d lineas (= transito) · %s · con mas de un pedido: %d", len(vista),
                     vista["vinculo"].value_counts().to_dict(), int(vista["n_pedidos"].gt(1).sum()))
    log_pedidos.info("=== FIN ===")


# ══════════════════════════ BACKFILL DESDE INBOUND SEMANAL ═══════════════════
def avisos_sin_linea(conocidos: set):
    """Avisos de llegada de los ultimos DIAS_AVISOS dias, sin respuestas. Devuelve los que no nombran ningun HAWB del bi
    ni del historial (uno por juego de IDs del asunto, con su ultimo correo), los HAWB conocidos que si nombran y los
    asuntos sin ningun ID."""
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    filtro = f"[ReceivedTime] >= '{datetime.now() - timedelta(days=DIAS_AVISOS):%m/%d/%Y %I:%M %p}'"
    pendientes, cubiertos, sin_id = {}, set(), set()
    for carpeta in C.carpetas_outlook(ns):
        for item in carpeta.Items.Restrict(filtro):
            try:
                asunto = getattr(item, "Subject", "") or ""
                if getattr(item, "Class", 0) != 43 or not C.ASUNTO_ARRIBO.search(asunto) or C.es_respuesta(asunto):
                    continue
                recibido = item.ReceivedTime.replace(tzinfo=None)
            except Exception as e:
                print(f"   correo omitido ({type(e).__name__})")
                continue
            ids = frozenset(k for k in C.ids_en_asunto(asunto) if sum(c.isdigit() for c in k) >= MIN_DIGITOS)
            if not ids:
                sin_id.add(asunto.strip())
            elif ids & conocidos:
                cubiertos |= ids & conocidos
            elif ids not in pendientes or recibido > pendientes[ids]["recibido"]:
                pendientes[ids] = {"recibido": recibido, "ids": ids, "tipo": "aviso"}
    return sorted(pendientes.values(), key=lambda a: a["recibido"]), cubiertos, sorted(sin_id)


def salidas_sin_linea(conocidos: set, avisos: list) -> list:
    """Salidas a DC de los ultimos DIAS_AVISOS dias que eventos ya registro (programaciones FWD de cualquier HAWB/BL)
    cuyo HAWB/BL no esta en el bi, ni en el historial, ni en un aviso de la lista: maritimos que salieron del
    in-transit antes de su salida a DC. Mismo formato que avisos_sin_linea."""
    if not C.RUTA_EVENTOS.exists():
        return []
    salidas = pd.read_excel(C.RUTA_EVENTOS, "salidas_dc", dtype={"hawb_key": str})
    salidas = salidas.assign(fecha_mail=pd.to_datetime(salidas["fecha_mail"], errors="coerce")).dropna(
        subset=["hawb_key", "fecha_mail"])
    ya = conocidos | set().union(*(a["ids"] for a in avisos))
    recientes = (salidas[salidas["fecha_mail"].ge(datetime.now() - timedelta(days=DIAS_AVISOS))
                         & ~salidas["hawb_key"].isin(ya)]
                 .sort_values("fecha_mail").drop_duplicates("hawb_key", keep="last"))
    return [{"recibido": f.to_pydatetime(), "ids": frozenset({k}), "tipo": "salida DC"}
            for k, f in zip(recientes["hawb_key"], recientes["fecha_mail"])]


def adjuntos_inbound(carpeta: Path) -> list:
    """Los Excel adjuntos (el Inbound y el Order Match) de los 'Inbound & Order Match' enviados en los ultimos
    DIAS_VENTANA_SALIDOS dias, del mas nuevo al mas viejo: (enviado, nombre, ruta)."""
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    desde = datetime.now() - timedelta(days=C.DIAS_VENTANA_SALIDOS)
    adjuntos = []
    for item in ns.GetDefaultFolder(5).Items.Restrict(f"[SentOn] >= '{desde:%m/%d/%Y %I:%M %p}'"):   # 5 = Enviados
        try:
            if ASUNTO_INBOUND not in (getattr(item, "Subject", "") or "").lower():
                continue
            enviado = pd.Timestamp(item.SentOn.replace(tzinfo=None)).floor("s")
            for adjunto in item.Attachments:
                nombre = str(adjunto.FileName)
                if nombre.lower().endswith(EXT_ADJUNTO):
                    ruta = carpeta / f"{enviado:%Y%m%d_%H%M%S}_{nombre}"
                    adjunto.SaveAsFile(str(ruta))
                    adjuntos.append((enviado, nombre, ruta))
        except Exception as e:
            print(f"   correo omitido ({type(e).__name__}: {e})")
    return sorted(adjuntos, key=lambda a: a[0], reverse=True)


def encabezados(crudo: pd.DataFrame, requeridas) -> tuple:
    """Fila de encabezados: la primera de las FILAS_ENCABEZADO que nombra todas las requeridas, como las nombra el job
    ('Transito[SO]' o 'SO' -> so). (None, None) si ninguna."""
    for i in range(min(FILAS_ENCABEZADO, len(crudo))):
        nombres = [limpiar_nombre(str(v)) for v in crudo.iloc[i]]
        if set(requeridas) <= set(nombres):
            return i, nombres
    return None, None


def pedidos_del_inbound(crudo: pd.DataFrame) -> pd.DataFrame:
    """Hoja Order Match (encabezados SO#, Material y EndCustomerName) -> so (422), mtm, so_local (433), cliente, sold_to.
    La 422 sale de 'SO Transit' y 'RPL SO' (sin ellas, de toda la fila); '4224382978-10' o 'a / b' se parten. Solo
    cuentan las que coinciden con una linea traida. Vacio si la hoja no es el Order Match."""
    vacio = pd.DataFrame(columns=["so", "mtm", "so_local", "cliente", "sold_to"])
    i, nombres = encabezados(crudo, {"so", "material"})
    cliente = next((c for c in CLIENTE_OM if nombres and c in nombres), None)
    if i is None or cliente is None:
        return vacio
    om = crudo.iloc[i + 1:].set_axis(nombres, axis=1)
    om = om.loc[:, ~om.columns.duplicated()]
    col_sold = columna_sold_to(om.columns)                                   # 'Sold_to'
    base = pd.DataFrame({"mtm": C.txt(om["material"]).str.upper(), "so_local": C.txt(om["so"]),
                         "cliente": om[cliente].astype("string").str.strip(),
                         "sold_to": C.txt(om[col_sold]) if col_sold else pd.NA}, index=om.index)
    cols_422 = [c for c in ("so_transit", "rpl_so") if c in om.columns]     # 'SO Transit' y 'RPL SO'
    origen = om[cols_422] if cols_422 else om.drop(columns=["so", "material", cliente])   # sin ellas: toda la fila
    piezas = origen.stack().astype(str).str.split(r"[^0-9A-Za-z]+", regex=True).explode()
    piezas = piezas[piezas.str.fullmatch(r"\d{6,}").fillna(False)]           # numeros de documento: la 422 esta entre ellos
    largo = base.loc[piezas.index.get_level_values(0)].assign(so=piezas.to_numpy())
    return largo.dropna(subset=["so", "so_local"]).drop_duplicates(["so", "mtm", "so_local"]).reset_index(drop=True)


def buscar(ruta: Path, ids: set):
    """Busca los ids por valor en cada celda de cada hoja. En una hoja de transito la fila donde esta el id es la linea,
    con el id como bl_awb_num (si la hoja tiene esa columna, solo cuenta ahi). Devuelve esas filas y en que otras hojas
    y columnas aparece cada id; de las hojas donde aparece sin ser de transito, sus 3 primeras filas (para ver sus
    encabezados reales); y los pedidos de su hoja Order Match."""
    filas, donde, primeras, pedidos = [], {}, {}, []
    with pd.ExcelFile(ruta) as libro:
        for hoja in libro.sheet_names:
            try:
                crudo = pd.read_excel(libro, hoja, header=None, dtype=str, keep_default_na=False, na_values=[""])
            except Exception as e:                       # hoja de grafico u otra cosa que no es tabla
                print(f"   {ruta.name} hoja '{hoja}' no se pudo leer ({type(e).__name__}): se salta")
                continue
            if crudo.empty:
                continue
            pedidos.append(pedidos_del_inbound(crudo))
            i, nombres = encabezados(crudo, COLS_PEDIDO)
            if i is None:                                # no es de transito: encabezados en la fila 1, solo para ubicar
                datos = crudo.iloc[1:].set_axis([limpiar_nombre(str(v)) for v in crudo.iloc[0]], axis=1)
            else:
                datos = crudo.iloc[i + 1:].set_axis(nombres, axis=1)
            datos = datos.loc[:, ~datos.columns.duplicated()]
            de_linea = set() if i is None else {"bl_awb_num"} if "bl_awb_num" in datos.columns else set(datos.columns)
            for col in datos.columns:
                valores = C.normalizar_alnum_serie(datos[col].dropna())
                hit = valores[valores.isin(ids)]
                if hit.empty:
                    continue
                if col in de_linea:
                    filas.append(datos.loc[hit.index].assign(bl_awb_num=datos.loc[hit.index, col], _k=hit,
                                                             _fila=hit.index + 1, _hoja=hoja))  # _fila: de Excel
                else:
                    for k in set(hit):
                        donde.setdefault(k, []).append(f"hoja '{hoja}' columna '{col}'")
                    if i is None:
                        primeras[hoja] = [crudo.iloc[r].dropna().astype(str).str[:25].tolist()[:25]
                                          for r in range(min(3, len(crudo)))]
    return ((pd.concat(filas, ignore_index=True) if filas else pd.DataFrame(columns=["_k"])), donde, primeras,
            pd.concat(pedidos, ignore_index=True).drop_duplicates(["so", "mtm", "so_local"]))


def fecha_excel(serie: pd.Series) -> pd.Series:
    """Serial de Excel (46000) o fecha en texto; lo ilegible queda vacio (en el historial la fecha no cruza nada)."""
    serial = pd.to_numeric(serie, errors="coerce")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)            # formatos mezclados: se lee cada valor
        texto = pd.to_datetime(serie.where(serial.isna()), errors="coerce")
    return pd.to_datetime(serial, unit="D", origin="1899-12-30", errors="coerce").fillna(texto)


def como_historial(filas: pd.DataFrame, enviado: pd.Timestamp) -> pd.DataFrame:
    """Filas de la hoja de transito con las columnas y tipos del historial."""
    lineas = filas.reindex(columns=HIST_TEXTO + [COL_QTY] + HIST_FECHAS)
    propia = "inbound-" + filas["_k"].astype(str) + "-" + filas["_fila"].astype(str)   # hoja sin Key: no choca con el bi
    lineas[COL_KEY] = lineas[COL_KEY].astype("string").fillna(propia.astype("string"))
    for c in (COL_KEY, COL_DN, COL_ITEM, "so"):
        lineas[c] = C.norm_doc(lineas[c])
    for c in HIST_FECHAS:
        lineas[c] = fecha_excel(lineas[c])
    lineas = lineas.assign(visto_desde=enviado, visto_hasta=enviado, fuente=f"inbound semanal {enviado:%d/%m/%Y}")
    return _tipar(lineas)[HIST_COLS]


def en_alcance_inbound(lineas: pd.DataFrame) -> pd.Series:
    """Misma regla que el bi: Peru, o Panama con el storage_loc de la rama que abastece Peru. Sin country, entra
    (conservada separada de cargar_reporte · R9 en PENDIENTE_UNIFICAR.md)."""
    pais = lineas["country"].astype("string").str.upper().fillna("")
    sloc = lineas["storage_loc"].astype("string").str.upper().fillna("")
    return (pais.eq("") | pais.eq(C.PAIS_PERU) | (pais.eq(C.PAIS_PANAMA) & sloc.isin(C.SLOC_RAMA_PANAMA))).astype(bool)


def pares(df: pd.DataFrame) -> pd.Series:
    """422 x MTM con la normalizacion del cruce del correo."""
    return C.txt(df["so"]) + "|" + C.txt(df["mtm"]).str.upper()


def previas_a_reintentar(rescatadas: pd.DataFrame, completos: set, hoy: pd.Timestamp) -> pd.DataFrame:
    """Lineas rescatadas antes que aun no tienen 433 o sold-to y vale la pena volver a buscar.
    Paso 4 (A4): una linea de stock nunca tendra 433; si su Inbound tiene mas de C.DIAS_REINTENTO_433 dias se deja
    de buscar. Antes, una sola linea de stock hacia releer los 30 dias de adjuntos en cada corrida."""
    sin_pedido = rescatadas[~pares(rescatadas).isin(completos)]
    recientes = pd.to_datetime(sin_pedido["visto_desde"], errors="coerce") >= hoy - pd.Timedelta(days=C.DIAS_REINTENTO_433)
    return sin_pedido[recientes.fillna(False)]


def rescatar(hist: pd.DataFrame):
    """Lo que el script agregaria, sin escribir nada: (lineas nuevas con las columnas del historial, una fila por aviso
    con lo que encontro, resumen con los pedidos nuevos). chequeo_cadena lo llama igual."""
    conocidos = set(C.clave_hawb(hist["bl_awb_num"]).dropna())
    avisos, cubiertos, sin_id = avisos_sin_linea(conocidos)
    salidas = salidas_sin_linea(conocidos, avisos)
    guardados = leer_historial_pedidos()
    con_pedido = set(pares(guardados).dropna())                                       # tienen 433
    completos = set(pares(guardados[guardados["sold_to"].notna()]).dropna())           # y sold-to
    rescatadas = hist[hist["fuente"].astype("string").str.startswith("inbound semanal").fillna(False)]
    previas = previas_a_reintentar(rescatadas, completos, pd.Timestamp.now())   # sin 433 o sin sold-to, recientes
    resumen = {"cubiertos": len(cubiertos), "sin_id": sin_id, "avisos": len(avisos), "salidas": len(salidas),
               "inbound": [], "primeras": {},
               "previas_sin_pedido": len(previas), "previas_con_pedido": 0, "previas_con_sold_to": 0,
               "pedidos": pd.DataFrame(columns=["so", "mtm", "so_local", "cliente", "sold_to", "enviado"]),
               "order_match": [], "sin_433": {}}
    avisos = avisos + salidas                           # desde aqui, cada correo con HAWB/BL que no esta
    if not avisos and previas.empty:
        return historial_vacio(), pd.DataFrame(columns=["correo", "tipo", "hawb", "resultado"]), resumen
    hallado, donde, oms = {}, {}, []                    # id -> (enviado, nombre, filas) | id -> [lugares] | Order Match
    carpeta = Path(tempfile.mkdtemp(prefix="backfill_transito_"))
    try:
        adjuntos = adjuntos_inbound(carpeta)
        resumen["inbound"] = [(enviado, nombre) for enviado, nombre, _ in adjuntos]
        for enviado, nombre, ruta in adjuntos:           # del mas nuevo al mas viejo: gana el mas nuevo
            faltan = set().union(*(a["ids"] for a in avisos if not a["ids"] & hallado.keys()))
            try:
                filas, en, primeras, pedidos = buscar(ruta, faltan)
            except Exception as e:
                print(f"   {nombre} no se pudo abrir ({type(e).__name__}: {e}): se salta")
                continue
            oms.append((enviado, pedidos))
            resumen["order_match"].append((nombre, len(pedidos)))
            for k, lugares in en.items():
                donde.setdefault(k, []).extend(f"{nombre} {lugar}" for lugar in lugares)
            resumen["primeras"].update({f"{nombre} hoja '{h}'": filas_h for h, filas_h in primeras.items()})
            for k, g in filas.groupby("_k"):
                if k not in hallado:
                    hallado[k] = (enviado, nombre, g)      # _k y _fila: llave propia si la hoja no trae Key
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)

    ya, nuevas, informe, sin_linea, entran_por_aviso = set(hist[COL_KEY].astype(str)), [], [], [], []
    for a in avisos:
        k = next((i for i in sorted(a["ids"]) if i in hallado), None)
        if k is None:
            lugar = next((lugar for i in sorted(a["ids"]) for lugar in donde.get(i, [])), None)
            texto = (f"esta en {lugar}, pero no como linea de transito (hoja con so y mtm, HAWB en bl_awb_num)" if lugar
                     else f"no esta en ninguna celda de los Inbound de {C.DIAS_VENTANA_SALIDOS} dias")
            informe.append([2 if lugar else 3, a["recibido"], a["tipo"], " / ".join(sorted(a["ids"])).upper(), texto])
            sin_linea += [lugar] if lugar else []
            continue
        enviado, nombre, filas = hallado[k]
        filas = filas[filas["_hoja"].eq(filas["_hoja"].iloc[0])]   # una sola hoja: el mismo HAWB en dos no se duplica
        lineas = como_historial(filas, enviado).drop_duplicates(COL_KEY)
        fuera, repetidas = ~en_alcance_inbound(lineas), lineas[COL_KEY].isin(ya)
        entran = lineas[~fuera & ~repetidas]
        texto = f"{nombre} ({enviado:%d/%m}) · {len(entran)} linea{'' if len(entran) == 1 else 's'}"
        if len(entran):
            texto += (f" · {', '.join(entran['mot'].dropna().unique())} · "
                      f"SO {', '.join(entran['so'].dropna().unique()[:3])}")
        if fuera.any():
            texto += f" · {int(fuera.sum())} fuera de alcance ({', '.join(lineas.loc[fuera, 'country'].dropna().unique())})"
        if repetidas.any():
            texto += f" · {int(repetidas.sum())} ya estaban en el historial"
        informe.append([1, a["recibido"], a["tipo"], filas["bl_awb_num"].iloc[0], texto])
        entran_por_aviso.append((len(informe) - 1, entran))
        nuevas.append(entran)
        ya |= set(entran[COL_KEY])                     # otro aviso del mismo HAWB no la repite
    nuevas = pd.concat(nuevas, ignore_index=True) if nuevas else historial_vacio()
    assert nuevas[COL_KEY].notna().all() and nuevas[COL_KEY].is_unique, "Key nula o repetida en lo que se agrega"

    # pedidos: los 422 x MTM de lo traido (ahora o antes) que historial_pedidos no tiene, del Order Match mas nuevo
    faltan_pares, tomados = (set(pares(nuevas).dropna()) | set(pares(previas).dropna())) - completos, []
    for enviado, pedidos in oms:
        if pedidos.empty or not faltan_pares:
            continue
        par = pares(pedidos)
        estos = pedidos[par.isin(faltan_pares).fillna(False)]
        if len(estos):
            tomados.append(estos.assign(enviado=enviado))
            faltan_pares -= set(par[par.isin(faltan_pares).fillna(False)])
    if tomados:
        resumen["pedidos"] = pd.concat(tomados, ignore_index=True)
    con_pedido |= set(pares(resumen["pedidos"]).dropna())
    completos |= set(pares(resumen["pedidos"][resumen["pedidos"]["sold_to"].notna()]).dropna())
    for fila, entran in entran_por_aviso:
        n = int(pares(entran).isin(con_pedido).sum())
        informe[fila][4] += f" · 433 en {n} de {len(entran)}" if len(entran) else ""
    resumen["previas_con_pedido"] = int(pares(previas).isin(con_pedido).sum())
    resumen["previas_con_sold_to"] = int(pares(previas).isin(completos).sum())
    sin_433 = pd.concat([nuevas, previas], ignore_index=True)
    sin_433 = sin_433[~pares(sin_433).isin(con_pedido)]
    so_en_om = set().union(*(set(C.txt(p["so"]).dropna()) for _, p in oms)) if oms else set()
    otro_material = C.txt(sin_433["so"]).isin(so_en_om)
    resumen["sin_433"] = {"lineas": len(sin_433), "otro_material": int(otro_material.sum()),
                          "ejemplo_otro": sin_433.loc[otro_material, ["bl_awb_num", "so", "mtm"]].head(3).values.tolist(),
                          "stock": sorted(sin_433.loc[~otro_material, "bl_awb_num"].dropna().unique())}

    resumen["primeras"] = {h: f for h, f in resumen["primeras"].items()          # solo donde quedo un HAWB sin linea
                           if any(lugar.startswith(h) for lugar in sin_linea)}
    informe = (pd.DataFrame(informe, columns=["orden", "correo", "tipo", "hawb", "resultado"])   # primero lo que entra
                 .sort_values(["orden", "correo"]).drop(columns="orden"))
    informe["correo"] = informe["correo"].map(lambda f: f"{f:%d/%m %H:%M}")
    return nuevas, informe, resumen


def main_backfill() -> None:
    hist = leer_historial_transito()
    assert len(hist), f"{C.HIST_PATH.name} vacio: correr antes el orquestador (job_in_transit lo arma con el bi de hoy)"
    print(f"historial: {hist[COL_KEY].nunique():,} lineas · {hist['bl_awb_num'].nunique():,} HAWB/BL · "
          f"{hist['visto_desde'].min():%d/%m %H:%M} -> {hist['visto_hasta'].max():%d/%m %H:%M}")
    nuevas, informe, r = rescatar(hist)
    print(f"avisos de llegada ({DIAS_AVISOS} dias): {r['cubiertos']} HAWB ya estan en el bi o el historial (los cruza "
          f"eventos) · {r['avisos']} avisos con HAWB que no esta · salidas a DC con HAWB/BL que no esta: {r['salidas']}")
    if r["sin_id"]:
        print(f"   {len(r['sin_id'])} avisos sin HAWB en el asunto (eventos tampoco los cruza): "
              + " | ".join(a[:60] for a in r["sin_id"][:5]))
    if r["previas_sin_pedido"]:
        print(f"lineas traidas en una corrida anterior sin pedido (433) o sin sold-to: {r['previas_sin_pedido']}")
    if not r["avisos"] and not r["salidas"] and not r["previas_sin_pedido"]:
        print("RESULTADO: nada que agregar — " + ("los avisos de esta semana ya tienen su HAWB en el bi o el historial"
              if r["cubiertos"] else f"no hay avisos de llegada con HAWB en los ultimos {DIAS_AVISOS} dias"))
        return
    print(f"Adjuntos de '{ASUNTO_INBOUND}' ({C.DIAS_VENTANA_SALIDOS} dias): "
          + (" · ".join(f"{n} ({e:%d/%m %H:%M})" for e, n in r["inbound"]) or "ninguno"))
    if not r["inbound"]:
        raise SystemExit(f"RESULTADO: NO se escribio nada — no hay '{ASUNTO_INBOUND}' con adjuntos Excel "
                         f"en Elementos enviados de los ultimos {C.DIAS_VENTANA_SALIDOS} dias")
    if len(informe):
        print("\n   correo       tipo       HAWB/BL         resultado")
        for correo, tipo, hawb, resultado in informe.itertuples(index=False):
            print(f"   {correo}  {tipo:<9}  {hawb:<15} {resultado}")
        print()
    pedidos = r["pedidos"]
    if len(pedidos) or r["previas_sin_pedido"]:
        print(f"pedidos (433) del Order Match de los Inbound: {len(pedidos)} para {pedidos['so'].nunique()} SO 422"
              + (f" · lineas de la corrida anterior que ahora tienen 433: {r['previas_con_pedido']} de "
                 f"{r['previas_sin_pedido']} · con sold-to: {r['previas_con_sold_to']}" if r["previas_sin_pedido"] else ""))
    print("Order Match en cada adjunto: " + (" · ".join(f"{n}: {k:,} 422 x Material -> 433" if k else f"{n}: no "
          "(sin hoja con SO#, Material y EndCustomerName)" for n, k in r["order_match"]) or "ninguno"))
    if r["order_match"] and not any(k for _, k in r["order_match"]):
        print("   ningun adjunto trae el Order Match: sin el, las lineas traidas quedan sin 433")
    s = r["sin_433"]
    if s.get("lineas"):
        print(f"   lineas traidas sin 433: {s['lineas']} · {s['otro_material']} con su 422 en el Order Match pero con otro "
              f"Material (ej. HAWB, 422, MTM: {s['ejemplo_otro']}) · el resto, su 422 no esta en ningun Order Match: "
              f"stock. El correo las muestra igual, como 'Sin pedido en OrderMatch'. HAWB/BL: {', '.join(s['stock'][:15])}")
    for hoja, filas_h in r["primeras"].items():           # donde esta el HAWB pero no se ve so ni mtm arriba
        print(f"{hoja}: el HAWB esta ahi pero sus primeras filas no nombran so y mtm —")
        for n, fila in enumerate(filas_h, 1):
            print(f"   fila {n}: {fila}")
    if nuevas.empty and pedidos.empty:
        if hist["fuente"].astype("string").str.startswith("inbound semanal").fillna(False).any():
            print("RESULTADO: nada nuevo — lo que estaba en los Inbound ya esta en el historial; los de la tabla no "
                  "estan como linea en ningun Inbound")
            return
        raise SystemExit("RESULTADO: NO se escribio nada — ni lineas ni pedidos nuevos en los Inbound "
                         "(la tabla de arriba dice donde esta cada HAWB)")
    hawbs = nuevas["bl_awb_num"].nunique()
    if C.DRY_RUN_BACKFILL:
        print(f"RESULTADO: DRY_RUN = True, NO se escribio nada — agregaria {len(nuevas)} lineas de {hawbs} HAWB y "
              f"{len(pedidos)} pedidos")
        return
    if len(nuevas):
        escribir_historial_transito(pd.concat([hist, nuevas], ignore_index=True))
        escrito = pq.read_metadata(C.HIST_PATH).num_rows      # solo el pie del parquet
        assert escrito == len(hist) + len(nuevas), (
            f"{C.HIST_PATH.name} quedo con {escrito} filas, se esperaban {len(hist) + len(nuevas)}")
    for enviado, grupo in pedidos.groupby("enviado"):
        guardar_pedidos(grupo, enviado, f"inbound semanal {enviado:%d/%m/%Y}")
    print(f"RESULTADO: OK — {len(nuevas)} lineas de {hawbs} HAWB a {C.HIST_PATH.name} · {len(pedidos)} pedidos (433) a "
          f"{C.HIST_PEDIDOS_PATH.name}\nahora corre el orquestador: eventos registra los avisos y el correo los muestra")
