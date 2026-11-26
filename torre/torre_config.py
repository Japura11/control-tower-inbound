# -*- coding: utf-8 -*-
"""
torre_config.py — Rutas, constantes compartidas, interruptores globales y helpers comunes
idénticos en comportamiento de la cadena Control Tower Perú (v2 · Paso 3).

Regla de arquitectura: este módulo NO importa ningún otro módulo del proyecto, ni abre COM,
Lake ni red. Todo lo que difiere sutilmente entre scripts (R5, R7 parcial, R9, R12 regex de aviso)
conserva sus variantes separadas y queda documentado en PENDIENTE_UNIFICAR.md para el Paso 4.
"""

import logging
import os
import re
import shutil
from datetime import datetime, timedelta  # noqa: F401 (usado por reloj virtual en pruebas)
from pathlib import Path

import numpy as np
import pandas as pd

# ══════════════════════════ RUTAS (una sola definición · R12) ════════════════
DIR_SCRIPT             = Path(__file__).resolve().parent
STAGING                = Path(os.environ.get("TORRE_STAGING", DIR_SCRIPT.parent / "sandbox"))   # produccion: definir TORRE_STAGING
CARPETA                = STAGING / "inbound_&_OM"
CARPETA_TRACKING       = STAGING / "Tracking"
CARPETA_M4             = STAGING / "deliverys_&_returns"

RUTA_OM                = CARPETA / "OM.xlsx"
RUTA_REPORTE           = RUTA_OM
RUTA_IN_TRANSIT        = CARPETA / "in_transit_peru.xlsx"
OUT_PATH               = RUTA_IN_TRANSIT
RUTA_EVENTOS           = CARPETA / "eventos_transito.xlsx"
RUTA_HISTORIAL         = CARPETA / "historial_transito.parquet"
HIST_PATH              = RUTA_HISTORIAL
RUTA_HISTORIAL_PEDIDOS = CARPETA / "historial_pedidos.parquet"
HIST_PEDIDOS_PATH      = RUTA_HISTORIAL_PEDIDOS

RUTA_SALIDA_HTML       = CARPETA_TRACKING / "radar_embarques.html"
RUTA_PLANTILLA         = DIR_SCRIPT / "radar_embarques_plantilla.html"
RUTA_CATALOGO          = DIR_SCRIPT / "catalogo_buques.json"      # IMO/MMSI y última posición de cada buque
RUTA_MAPA              = DIR_SCRIPT / "mapa_mundial_50m.json"     # geometría del mundo: se baja una sola vez

# ══════════════════════════ INTERRUPTORES GLOBALES ═══════════════════════════
USAR_LAKE        = os.environ.get("TORRE_USAR_LAKE", "1") != "0"   # paso 5: python torre.py diario --sin-lake (sin bultos/kg/m3 de Lake)
DRY_RUN_EVENTOS  = False           # True: lee, resuelve y muestra lo que cambiaría; no escribe historial ni hojas
DRY_RUN_BACKFILL = False           # True: muestra lo que agregaría y no escribe
DRY_RUN_CORREO   = True            # True: borrador abierto en Outlook · False: se envía
USAR_AIS         = True            # True: consulta ProveedorA + ProveedorB con caché y freno
ABRIR_NAVEGADOR  = False           # True: abre el radar al terminar (corrida a mano)

# ══════════════════════════ HOJAS Y COLUMNAS COMPARTIDAS (R12) ═══════════════
HOJA_BI          = "bi"
HOJA_DN_VOLUMEN  = "dn_volumen"
HOJA_HU          = "hu_detail"
HOJA_SEGUIMIENTO = "seguimiento"
HOJA_VISTA       = HOJA_SEGUIMIENTO
HOJA_EMBARQUES   = "embarques"
HOJA_PEDIDOS     = "pedidos_locales"
HOJA_ARRIBOS     = "arribos"
HOJA_SALIDAS_DC  = "salidas_dc"
HOJA_NOTIF       = "notificaciones"

HOJAS_AJENAS_EVENTOS      = {"bi", "dn_volumen", "hu_detail", "meta"}   # de ingesta (in_transit): nunca se escriben desde eventos
HOJAS_AGUAS_ABAJO_EVENTOS = ["pedidos_locales"]                         # se arman desde seguimiento: al rehacerla se retiran
HOJAS_AJENAS_PEDIDOS      = {"bi", "dn_volumen", "hu_detail", "meta", "seguimiento", "embarques"}

COL_KEY      = "key"                   # llave unica del reporte (grain de bi)
COL_DN       = "delivery_order_number" # llave contra Lake.delivery_number
COL_ITEM     = "delivery_item_number"
COL_QTY      = "transit_qty"
COL_HAWB     = "bl_awb_num"
COL_SO       = "so"
COL_MOT      = "mot"
COL_ATA      = "ata"                   # llegada a puerto segun el reporte: la unica senal de arribo en maritimos
COL_CORTE_BI = "corte_bi"              # hora en que job_in_transit escribio el bi; viaja en la vista

# ══════════════════════════ REGLAS Y CONSTANTES DE NEGOCIO (R12) ═════════════
# Alcance geografico = regla validada del job anterior: Peru + rama Panama que abastece Peru
PAIS_PERU        = "PE"
PAIS_PANAMA      = "PA"
SLOC_RAMA_PANAMA = ["PE01", "WBPE"]

MOT_MARITIMO         = {"SEA", "SEA-AIR"}  # arriban por ATA o ETA vencida de los correos FWD, no por el aviso
MOT_AEREO            = "AIR"
MAX_HORAS            = 12                  # bi / seguimiento / OM.xlsx mas viejo que esto -> aborta
HORAS_CACHE_AIS_RADAR = 2                  # paso 5: el radar horario re-consulta la posicion si tiene mas de esto
CARPETA_FOTOS        = STAGING / "fotos_semanales"   # paso 5: copia semanal fija de los historiales (nunca se reescribe)
MAX_HORAS_RADAR      = 24                  # paso 5: el radar horario trabaja sobre la base del dia (corrida diaria)
DIAS_VENTANA_SALIDOS = 30                  # ventana de 30 dias para historial / salidos / Inbound semanal
DIAS_HABILES_ENTREGA = 2                   # sin recepcion 3PL: entregado al DC = salida FWD + 2 dias habiles (lun-vie)

ORDEN_HITO      = ["En tránsito internacional", "Arribado", "En tránsito a DC", "Entregado al DC"]
ARRIBADO        = "Arribado"
EN_RUTA         = "En ruta al DC"
AVISOS_PASO: list = []              # paso 4: lo que un paso no pudo hacer pero no detiene la cadena (sale con codigo 2)
SALIDA_CON_AVISO = 2
M4_ARCHIVOS = ("m4_latam.parquet", "m4_latam_sin_gi.parquet")   # paso 5: ingreso en SAP = 433 con DN (diag_ingreso_dc)
PAIS_M4 = "PERU"
UMBRAL_ALERTA_ETA_DIAS = 5             # paso 5: alerta si la ETA vigente se mueve mas de esto (usuario 25/11)
DIAS_REINTENTO_433 = 7                 # paso 4 (A4): una linea rescatada sin 433 se re-busca una semana (luego es stock)
REINTENTOS_REEMPLAZO, PAUSA_REEMPLAZO_S = 3, 0.5   # paso 4 (A3): locks breves de OneDrive

SLA_ARRIBO_HAB  = {"AIR": 3, "SEA": 7, "SEA-AIR": 7}   # dias habiles desde la llegada sin salida a DC (decidido 23/10)
PASO_FUERA_SLA  = "pedir a FWD la programación de salida"
PASO_AVISO      = "registrar el reclamo con el agente de carga"

COLS_AVISO          = ["aviso_total", "aviso_buenos", "aviso_malos", "aviso_faltantes"]
COLS_EMBARQUE       = ["n_bultos", "peso_bruto_kg", "volumen_m3",
                       "aviso_total", "aviso_buenos", "aviso_malos", "aviso_faltantes"]
COLS_EMBARQUE_RADAR = ["n_bultos", "peso_bruto_kg", "volumen_m3"]

# delivery_deleted: dominio cerrado. Un valor nuevo revienta en job_in_transit en vez de sumarse como vigente
ELIMINADA_SI = {"X", "Y", "YES", "TRUE", "1"}
ELIMINADA_NO = {"N", "NO", "FALSE", "0"}

# hu_detail trae Lake tal cual. Las vistas convierten a kg / cm / m3 con la unidad de CADA bulto.
# Lake usa codigos ISO (UN/ECE Rec. 20: CMT, CMQ, KGM); se dejan tambien los internos de SAP (CM, CCM, KG)
FACTOR_PESO_KG = {"KGM": 1.0, "GRM": 0.001, "LBR": 0.45359237, "TNE": 1000.0, "ONZ": 0.028349523,
                  "KG": 1.0, "G": 0.001, "LB": 0.45359237, "TO": 1000.0}
FACTOR_DIM_CM  = {"CMT": 1.0, "MMT": 0.1, "DMT": 10.0, "MTR": 100.0, "INH": 2.54, "FOT": 30.48,
                  "CM": 1.0, "MM": 0.1, "M": 100.0, "IN": 2.54}
FACTOR_VOL_M3  = {"CMQ": 1e-6, "MMQ": 1e-9, "DMQ": 1e-3, "LTR": 1e-3, "MTQ": 1.0, "FTQ": 0.0283168466,
                  "INQ": 1.6387064e-5, "CCM": 1e-6, "CDM": 1e-3, "L": 1e-3, "M3": 1.0, "CM3": 1e-6}
#               medida                  valor crudo          unidad               factores
CONVERSIONES = {"largo_cm":         ("length",            "unit_of_dimension", FACTOR_DIM_CM),
                "ancho_cm":         ("width",             "unit_of_dimension", FACTOR_DIM_CM),
                "alto_cm":          ("height",            "unit_of_dimension", FACTOR_DIM_CM),
                "volumen_m3":       ("total_volume",      "volume_unit",       FACTOR_VOL_M3),
                "peso_bruto_kg":    ("total_weight",      "weight_unit",       FACTOR_PESO_KG),
                "peso_cobrable_kg": ("chargeable_weight", "weight_unit",       FACTOR_PESO_KG)}

# Buzones y patrones de asuntos de Outlook compartidos entre eventos, backfill y chequeo
STORES_BUZON  = ("operaciones", "cliente")                   # stores de Outlook que se recorren
CARPETA_EXTRA = "Shipments"
LARGO_MIN_ID  = 6                                       # llaves mas cortas son ruido (regla del flujo validado)
ASUNTO_ARRIBO = re.compile(r"aviso\s+de\s+llegada", re.I)
RE_RESPUESTA  = re.compile(r"^(RE|RV|FW|FWD|R)\s*:", re.I)
RE_ETIQUETAS  = re.compile(r"\[External\]|\[#TN#[^\]]*\]", re.I)
RE_ID_ASUNTO  = re.compile(r"[A-Z0-9]{3,}[-\s]?\d{6,}") # prefijo-numero: 157-12345678, SHA1234567
RE_TOKEN      = re.compile(r"[A-Z0-9]+")

# Literales nulos separados (R5: NO unificados en Paso 3 porque difieren en "-" y "N/A")
NULOS_TEXTO_BI = ["NONE", "NAN", "NULL", "NAT", "<NA>", "(BLANK)", "(EN BLANCO)"]
NULOS_PEDIDOS  = {"", "NAN", "NONE", "NULL", "NAT", "<NA>", "(BLANK)", "(EN BLANCO)", "-", "N/A"}


# ══════════════════════════ NORMALIZACIÓN DE IDS (R5: variantes preservadas) ═
def es_vacio_bi(serie: pd.Series) -> pd.Series:
    """Nulidad sobre la serie CRUDA en job_in_transit; la lista de literales es solo red secundaria (sin '-' ni 'N/A')."""
    texto = serie.astype("string").str.strip()
    return serie.isna() | texto.eq("") | texto.str.upper().isin(NULOS_TEXTO_BI)


def norm_doc(serie: pd.Series) -> pd.Series:
    """Numero de documento como texto: sin espacios y sin el '.0' que Excel agrega al leer un numero.
    Idéntica byte a byte a job_in_transit.norm_doc y job_eventos_transito.texto_id."""
    return serie.astype("string").str.strip().str.replace(r"\.0+$", "", regex=True)


def txt(serie: pd.Series) -> pd.Series:
    """Normalización de job_pedidos_transito: norm_doc + reemplazo a <NA> de NULOS_PEDIDOS (incluye '-' y 'N/A')."""
    s = serie.astype("string").str.strip().str.replace(r"\.0+$", "", regex=True)
    return s.where(~s.str.upper().isin(NULOS_PEDIDOS))


def normalizar_id(texto) -> str:
    """HAWB/BL comparable entre fuentes: sin el '.0' de Excel, solo alfanumerico, minusculas."""
    base = re.sub(r"\.0+$", "", str(texto).strip())
    return re.sub(r"[^A-Za-z0-9]", "", base).lower()


def clave_hawb(serie: pd.Series) -> pd.Series:
    """Misma regla que normalizar_id sobre una columna (map: pocas centenas de filas), exigiendo len >= LARGO_MIN_ID."""
    clave = serie.dropna().astype(str).map(normalizar_id)
    return clave.where(clave.str.len() >= LARGO_MIN_ID).reindex(serie.index)


def normalizar_alnum_serie(serie: pd.Series) -> pd.Series:
    """normalizar_id vectorizada sobre una columna entera sin tope de largo (backfill: un Inbound tiene decenas de miles de celdas)."""
    return (serie.astype("string").str.strip().str.replace(r"\.0+$", "", regex=True)
                 .str.replace(r"[^A-Za-z0-9]", "", regex=True).str.lower())


# ══════════════════════════ CALENDARIO Y FECHAS (R7) ═════════════════════════
def edad_horas(momento, ahora=None) -> float:
    """Antigüedad en horas respecto de `ahora` (por defecto pd.Timestamp.now()); inf si momento es vacío."""
    if momento is None or momento == "" or pd.isna(momento):
        return float("inf")
    ref = pd.Timestamp.now() if ahora is None else ahora
    return (ref - pd.Timestamp(momento)).total_seconds() / 3600


def dias_habiles(desde: pd.Series, hasta: pd.Timestamp) -> pd.Series:
    """Dias habiles lun-vie entre cada fecha y `hasta` (sin feriados; versión job_eventos_transito sin clip)."""
    ok = desde.notna()
    out = pd.Series(np.nan, index=desde.index)
    out[ok] = np.busday_count(desde[ok].values.astype("datetime64[D]"), np.datetime64(hasta.date(), "D"))
    return out


def dias_habiles_no_neg(desde: pd.Series, hasta: pd.Timestamp) -> pd.Series:
    """Dias habiles lun-vie desde cada fecha hasta `hasta` truncando negativos a 0 (versión job_mail_arribados)."""
    return dias_habiles(desde, hasta).clip(lower=0)


# ══════════════════════════ ESCRITURA ATÓMICA (R10) ══════════════════════════
def reemplazar_atomico(tmp: Path, destino: Path) -> None:
    """Reemplazo atómico de archivo temporal sobre destino con mensaje claro ante bloqueo.
    Paso 4 (A3): OneDrive toma locks de segundos tras cerrar un archivo; se reintenta antes de rendirse."""
    import time
    for intento in range(REINTENTOS_REEMPLAZO):
        try:
            os.replace(tmp, destino)
            return
        except PermissionError as e:
            if intento == REINTENTOS_REEMPLAZO - 1:
                raise RuntimeError(f"{destino.name} esta abierto: cerrarlo y re-correr "
                                   f"(lo nuevo quedo en {tmp.name})") from e
            time.sleep(PAUSA_REEMPLAZO_S)


def reemplazar_hojas_excel(ruta: Path, hojas: dict, mtime_leido: float,
                           hojas_ajenas: set, retirar=(), logger=None) -> None:
    """Reemplaza atómicamente solo las hojas indicadas en `ruta`, verificando mtime y retirando derivadas si aplica."""
    ajenas = set(hojas) & set(hojas_ajenas)
    assert not ajenas, f"{sorted(ajenas)} son hojas de otro job"
    if ruta.stat().st_mtime != mtime_leido:
        raise RuntimeError(f"{ruta.name} cambio mientras este job corria — re-correr")
    tmp = ruta.with_name(ruta.stem + ".tmp.xlsx")
    shutil.copy2(ruta, tmp)
    with pd.ExcelWriter(tmp, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
        for derivada in retirar:
            if derivada in w.book.sheetnames:
                w.book.remove(w.book[derivada])
                if logger is not None:
                    logger.info("   %s retirada: se arma desde seguimiento — correr job_pedidos_transito", derivada)
    reemplazar_atomico(tmp, ruta)


# ══════════════════════════ LECTURAS Y CONVERSIONES COMPARTIDAS (R3) ═════════
def sold_to_por_pedido(avisar_sin_columna: bool = True, logger=None) -> dict:
    """433 -> sold-to, de historial_pedidos.parquet: la copia de hoy de job_in_transit y lo traido de los Inbound. El
    sold-to es del pedido (uno por 433): manda el ultimo visto. Sin la columna (copia de antes del 19/11): vacio."""
    ruta = RUTA_HISTORIAL_PEDIDOS
    if not ruta.exists():
        return {}
    guardados = pd.read_parquet(ruta)
    if "sold_to" not in guardados.columns:
        if avisar_sin_columna and logger is not None:
            logger.warning("%s sin sold_to: correr job_in_transit actualizado (lo agrega en su copia)", ruta.name)
        return {}
    guardados = (guardados.assign(so_local=txt(guardados["so_local"]), sold_to=txt(guardados["sold_to"]))
                          .dropna(subset=["so_local", "sold_to"]).sort_values("visto_hasta"))
    return guardados.drop_duplicates("so_local", keep="last").set_index("so_local")["sold_to"].to_dict()


def leer_embarques(columnas=COLS_EMBARQUE, deduplicar: bool = False, logger=None) -> pd.DataFrame:
    """Hoja embarques de job_eventos_transito: bultos, kg y m3 del HAWB/BL completo, ya convertidos (crudo de Lake)."""
    cols = list(columnas)
    with pd.ExcelFile(RUTA_IN_TRANSIT) as libro:
        if HOJA_EMBARQUES not in libro.sheet_names:
            if logger is not None:
                logger.warning("%s sin hoja %r: sale sin bultos ni volumen", RUTA_IN_TRANSIT.name, HOJA_EMBARQUES)
            return pd.DataFrame(columns=cols)
        emb = pd.read_excel(libro, HOJA_EMBARQUES, dtype={"embarque": str})
    limpio = emb.assign(embarque=txt(emb["embarque"])).dropna(subset=["embarque"])
    if deduplicar:
        limpio = limpio.drop_duplicates("embarque")
    return limpio.set_index("embarque").reindex(columns=cols)


def en_unidades_estandar(hu: pd.DataFrame) -> pd.DataFrame:
    """Bultos vigentes con sus medidas en kg / cm / m3, convertidas con la unidad de cada bulto.
    Un codigo sin factor deja esa medida vacia y avisa: la vista sale igual, sin inventar la unidad."""
    log_ev = logging.getLogger("eventos_transito")
    borrada = hu["delivery_deleted"].astype("string").str.strip().str.upper().isin(ELIMINADA_SI).fillna(False)
    v = hu[~borrada.astype(bool)].copy()
    for destino, (crudo, unidad, factores) in CONVERSIONES.items():
        codigo = v[unidad].astype("string").str.strip().str.upper()
        sin_factor = sorted(set(codigo[pd.to_numeric(v[crudo], errors="coerce").fillna(0).ne(0)].dropna())
                            - set(factores))
        if sin_factor:
            log_ev.warning("   %s: unidades sin factor %s — esos bultos quedan sin esa medida; agregar el factor",
                           destino, sin_factor)
        v[destino] = pd.to_numeric(v[crudo], errors="coerce") * pd.to_numeric(codigo.map(factores), errors="coerce")
    return v


# ══════════════════════════ HELPERS PUROS DE OUTLOOK ═════════════════════════
def es_respuesta(asunto) -> bool:
    return bool(RE_RESPUESTA.match(re.sub(r"\[.*?\]", "", str(asunto)).strip()))


def ids_en_asunto(asunto) -> set:
    """Candidatos a HAWB/BL en el asunto. El patron prefijo-numero se pega a la palabra previa
    ('HAWB 1234567890' -> 'hawb1234567890'); cada token alfanumerico suelto lo rescata."""
    limpio = RE_ETIQUETAS.sub("", str(asunto)).upper()
    candidatos = RE_ID_ASUNTO.findall(limpio) + RE_TOKEN.findall(limpio)
    return {k for k in map(normalizar_id, candidatos) if len(k) >= LARGO_MIN_ID}


def carpetas_outlook(ns):
    log_ev = logging.getLogger("eventos_transito")
    for store in ns.Stores:
        if not any(t in store.DisplayName.lower() for t in STORES_BUZON):
            continue
        try:
            yield store.GetDefaultFolder(6)             # 6 = olFolderInbox
        except Exception as e:
            log_ev.warning("Inbox inaccesible en %s: %s", store.DisplayName, e)
        try:
            yield store.GetRootFolder().Folders[CARPETA_EXTRA]
        except Exception:                               # no todos los stores tienen la carpeta
            pass
