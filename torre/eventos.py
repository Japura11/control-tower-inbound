# -*- coding: utf-8 -*-
"""
eventos.py — Consolidación sin cambios de lógica de:
  1. correos_fwd.py           (lectura pura de correos FWD — prealerta, notificación, aviso — y recepción 3PL)
  2. job_eventos_transito.py (escaneo de Outlook, historial append-only en eventos_transito.xlsx y hojas
                              seguimiento / embarques en in_transit_peru.xlsx) -> main_eventos()

Reglas validadas el 17/11/2026 contra 16 BL marítimos y el reporte Transito (ATD a ±1 día: 12 de 12):
  Fechas  Año primero (2026-09-08, 2026/09/08) = año-mes-día; con el mes en letras, como venga ('17 nov 2026',
          'Nov 17, 2026'); el resto, día/mes: 70 de 70 bloques con fechas no ambiguas eran día/mes. Sin año: el más
          cercano al correo. A más de 150 días del correo: ruido (dirección, contenedor).
  Bloques asunto, mensaje de arriba, mensajes citados del hilo y cada PDF. El hilo se corta en la cabecera de
          Outlook (From/De + Sent/Enviado), en '-----Original Message-----', en la línea de guiones bajos, en
          'On … wrote:' / 'El … escribió:' y en las líneas que empiezan con '>'. En un correo nuevo gana el PDF; en
          una respuesta gana lo escrito arriba y el asunto no cuenta (arrastra la fecha del original).
  Correo  ETA = New ETA > ETA/ATA futura > ETA, solo de texto propio: el citado arrastra la ETA del original.
          ATA = ATA o ETA/ATA no posterior al correo. ATD = ATD o ETD/ATD no posterior; también del citado.
  BL      ETA vigente = la del aviso, notificación o respuesta más reciente; la prealerta original o reenviada,
          solo si no hay ninguno de esos.
          ATA = la más reciente, salvo que una ETA de ese correo o de uno posterior caiga más de 2 días después
          (era una ETA vieja, no un arribo). ATD = la más reciente. Recepción 3PL = la primera.
"""

import logging
import os
import re
import shutil
import tempfile
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import pdfplumber                      # type: ignore
import win32com.client as win32        # type: ignore
from bs4 import BeautifulSoup
from dateutil import parser as dparser

import torre_config as C

# ══════════════════════════ CONFIG · CORREOS FWD ══════════════════════════════
TIPOS = [("prealerta", re.compile(r"pre-?\s*alert|prealerta", re.I)),
         ("notificacion", re.compile(r"cliente\s+shipment\s*[/|\-]", re.I)),
         ("aviso", re.compile(r"aviso\s+de\s+llegada|arrival\s+notice", re.I)),
         ("recepcion_dc", re.compile(r"importaci[oó]n\s+cliente\s*-\s*recepci[oó]n", re.I))]
TIPOS_CON_PDF = {"prealerta", "aviso"}
TIPOS_PRINCIPALES = {"aviso", "notificacion"}      # su ETA manda sobre la de la prealerta
MAX_DESFASE_DIAS = 150
DIAS_ETA_ANULA_ATA = 2

ES_A_EN = {"ene": "jan", "feb": "feb", "mar": "mar", "abr": "apr", "may": "may", "jun": "jun", "jul": "jul",
           "ago": "aug", "sep": "sep", "set": "sep", "oct": "oct", "nov": "nov", "dic": "dec"}
MESES = "ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic|jan|apr|aug|dec"
FECHA = (rf"\d{{4}}[-/.]\d{{1,2}}[-/.]\d{{1,2}}|\d{{1,2}}[/.-]\d{{1,2}}(?:[/.-]\d{{2,4}})?"
         rf"|\d{{1,2}}(?:\s+de)?[\s-]*(?:{MESES})[a-z]*\.?(?:(?:\s+de)?[\s-]*\d{{2,4}})?"
         rf"|(?:{MESES})[a-z]*\.?\s+\d{{1,2}}(?:st|nd|rd|th)?,?(?:\s+\d{{4}})?")
RE_FECHA = re.compile(FECHA, re.I)
RE_ISO = re.compile(r"^\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")
# orden = prioridad dentro de una línea: lo específico primero; cada calce se borra para que 'eta' no relea 'new eta'
ETA, ATA, ATD = r"e\.?t\.?a\.?", r"a\.?t\.?a\.?", r"a\.?t\.?d\.?"      # ETA y E.T.A.
CAMPOS = {"new_eta": rf"(?:new|nueva|revised|updated)\s+{ETA}|{ETA}\s+(?:actualizada|nueva|revisada)",
          "previous_eta": rf"(?:previous|old|anterior)\s+{ETA}|{ETA}\s+anterior",
          "eta_ata": rf"{ETA}\s*/\s*{ATA}", "etd_atd": r"e\.?t\.?d\.?\s*/\s*" + ATD,
          "ata": rf"(?<![\w/]){ATA}(?![\w/])", "atd": rf"(?<![\w/]){ATD}(?![\w/])", "eta": rf"(?<![\w/]){ETA}(?![\w/])"}
RE_CAMPO = {c: re.compile(rf"(?:{e})\s*[:\-.]?\s*({FECHA})", re.I) for c, e in CAMPOS.items()}
RANGO_BLOQUE = {False: {"asunto": 0, "cuerpo": 1, "pdf": 2}, True: {"pdf": 0, "cuerpo": 1}}   # gana el mayor
RE_CAB_DE = re.compile(r"^\s*(?:from|de)\s*:", re.I)
RE_CAB_ENVIO = re.compile(r"^\s*(?:sent|enviado(?:\s+el)?|date|fecha)\s*:", re.I)
RE_SEPARADOR = re.compile(r"^\s*(?:-{2,}\s*(?:original message|mensaje original|forwarded message|mensaje reenviado)"
                          r"\s*-{2,}|_{8,}\s*$|(?:on|el)\b.{5,150}\b(?:wrote|escribi[oó])\s*:\s*$)", re.I)
RE_CITA_MAYOR = re.compile(r"^\s*>")
RE_PUERTO = re.compile(r"(?:p\.?\s*of\s*discharge|pto\.?\s*(?:de\s*)?descarga)\s*[:\-]?\s*"
                       r"([A-Za-z][A-Za-z ]{2,20}?)(?=\s{2,}|\s+et[ad]\b|\s*$|,|/)", re.I | re.M)
RE_LOCODE_PE = re.compile(r"\bPE(CLL|CHY|PAI|MRI|ILQ)\b")
LOCODE_PE = {"CLL": "CALLAO", "CHY": "CHANCAY", "PAI": "PAITA", "MRI": "MATARANI", "ILQ": "ILO"}
RE_POL = re.compile(r"\bpol\s*:\s*([A-Z]{5})\b", re.I)
RE_BUQUE = re.compile(r"\bbuque\s*:\s*([A-Z0-9][A-Z0-9 .\-]{2,30}?)\s*(?://|$)", re.I)
COLS_RESUELTO = ["eta", "eta_fuente", "eta_tipo", "ata", "ata_fuente", "ata_anulada", "atd", "atd_fuente",
                 "recepcion_dc", "puerto", "pol", "buque"]

# ══════════════════════════ CONFIG · EVENTOS DE TRÁNSITO ═════════════════════
SIN_HAWB   = "(sin HAWB)"
PERFIL_TOP = 3                     # medidas de bulto que se nombran en perfil_bultos; el resto va como "+N"

COLS_HU_CRUDAS = (["delivery_number", "delivery_deleted"]
                  + list(dict.fromkeys(c for par in C.CONVERSIONES.values() for c in par[:2])))

COL_KEY, COL_HAWB, COL_SO, COL_MOT = C.COL_KEY, C.COL_HAWB, C.COL_SO, C.COL_MOT
COL_ATA      = C.COL_ATA
COL_CORTE_BI = C.COL_CORTE_BI
COLS_ID_BI   = [COL_KEY, COL_HAWB, COL_SO, COL_MOT, "delivery_order_number", "delivery_item_number",
                "dc_sto", "po", "bill"]   # como texto: Key tiene 16 digitos y como float pierde precision

# Orden de la vista: identificacion, veredicto y evidencia primero; trazabilidad al final
COLS_FRENTE = [COL_SO, "delivery_order_number", COL_HAWB, COL_MOT, "hito",
               "llegada_fecha", "llegada_fecha_fuente", "salida_dc_fecha", "salida_dc_hora",
               "recepcion_dc", "eta_correo", "atd_correo", "puerto_descarga"]
COLS_TRAZA  = ["llegada_aviso_mail", "salida_dc_mail", COL_CORTE_BI]

DIAS_VENTANA_INICIAL = 30      # primera corrida, con historial vacio
DIAS_VENTANA_FWD      = 60      # primera corrida de los correos FWD (hoja notificaciones vacia)
MARGEN_DIAS          = 2       # solape sobre el ultimo mail registrado; el append es idempotente

ASUNTO_ENVIO           = "PROGRAMACION DE ENTREGA CLIENTE"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

COLS_BULTOS    = ["bultos_total", "bultos_buenos", "bultos_malos", "bultos_faltantes"]
COLS_ARRIBOS   = ["hawb", "hawb_key", "id_mail", "fecha_mail", "fecha_llegada",
                  *COLS_BULTOS, "so", "mot", "capturado"]   # un historial viejo con 'comentario': se ignora
NUEVAS_ARRIBOS = set(COLS_BULTOS)  # agregadas despues de la primera version: un historial viejo las recibe vacias
COLS_SALIDAS   = ["hawb", "hawb_key", "id_mail", "fecha_mail", "fecha_programada",
                  "hora_estimada", "so", "mot", "capturado"]
COLS_NOTIF     = ["hawb", "hawb_key", "id_mail", "fecha_mail", "tipo", "respuesta", "eta", "eta_fuente",
                  "ata", "ata_fuente", "atd", "atd_fuente", "recepcion_dc", "puerto", "pol", "buque",
                  "so", "mot", "capturado"]
COLS_TEXTO_HIST = {"hawb", "hawb_key", "id_mail", "hora_estimada", "so", "mot",
                   "tipo", "eta_fuente", "ata_fuente", "atd_fuente", "puerto", "pol", "buque"}
COLS_FECHA_HIST = {"fecha_mail", "fecha_llegada", "fecha_programada", "capturado", "eta", "ata", "atd", "recepcion_dc"}

# ── Reglas de parseo validadas en produccion (vienen de rpa_maestro) ─────────
RE_FECHA_PROG = re.compile(r"programaci[oó]n\s+para\s+el\s+(?:d[ií]a\s+)?(\d{1,2}[/-]\d{1,2})", re.I)
RE_HORA_PROG  = re.compile(r"(\d{1,2}:\d{2})\s*Hrs", re.I)
PATRONES_LLEGADA = [re.compile(p, re.I) for p in (          # en orden de preferencia
    r"Fecha\s+de\s+llegada\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
    r"Fecha\s+de\s+arribo\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
    r"Fecha\s+de\s+ingreso\s+al\s+Dep\.?\s+Temp\.?\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
    r"ETA\s*/\s*ATA\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
    r"Ingreso\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
)]
RE_FECHA_SUELTA     = re.compile(r"\d{2}/\d{2}/\d{4}")       # para mostrar donde esta la fecha si ningun patron calza
RE_BULTOS_FILA      = re.compile(r"Bultos?\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)")
RE_BULTOS_ETIQUETAS = re.compile(r"Manifestados.*?(\d+).*?Buenos.*?(\d+).*?Malos.*?(\d+)", re.DOTALL)

ALIAS_COLUMNAS = {
    "ORDEN": ["ORDEN", "ORDER", "N°", "NRO"],
    "FILE": ["FILE", "FILE:", "ARCHIVO"],
    "BL": ["BL", "B/L", "BL.", "BILL OF LADING"],
    "FACTURA": ["FACTURA", "INVOICE", "FACT"],
    "TERMINAL": ["TERMINAL", "TERM"],
    "BULT": ["BULT", "BULTOS", "PKGS", "PIEZAS"],
    "CANTIDAD": ["CANTIDAD", "QTY", "CTDAD"],
    "PESO": ["PESO", "WEIGHT", "KGS", "KG"],
    "OBSERVACIONES": ["OBSERVACIONES", "OBS", "REMARKS", "NOTAS"],
    "ENTREGA": ["ENTREGA", "DELIVERY", "FECHA ENTREGA"],
    "DESTINO": ["DESTINO", "DESTINATION", "DEST"],
    "CLIENTE": ["CLIENTE", "CLIENT", "CONSIGNEE"],
    "AWB": ["AWB", "AIR WAYBILL", "HAWB"],
}
_ALIAS_LOOKUP   = {v.upper().strip(): canon for canon, vs in ALIAS_COLUMNAS.items() for v in vs}
HEADERS_VALIDOS = {"ORDEN", "CLIENTE", "AWB", "FACTURA", "TERMINAL", "BULT", "BULTOS",
                   "CANTIDAD", "PESO", "OBSERVACIONES", "ENTREGA", "DESTINO", "FILE", "BL"}
STOPWORDS_CORTE = {"SALUDOS", "BEST REGARDS", "REGARDS", "ATENTAMENTE", "CORDIALMENTE",
                   "DATOS DE TRANSPORTE", "FURGON", "BREVETE"}
ORDEN_PATTERN   = re.compile(r"^\d{1,6}$")
RE_FACTURA      = re.compile(r"^\d{7,12}$")

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("eventos_transito")


# ══════════════════════════ FUNCIONES PURAS · CORREOS FWD ═════════════════════
def tipo_de(asunto) -> str | None:
    for nombre, patron in TIPOS:
        if patron.search(str(asunto)):
            return nombre
    return None


def a_fecha(texto: str, ref: datetime):
    """ISO = año-mes-día; lo demás día/mes ('17/11', '25.10.2026', '17 de noviembre'). Sin año: el más cercano al
    correo. Fuera de ±150 días del correo: None."""
    m = RE_ISO.match(texto)
    if m:
        try:
            f = datetime(*map(int, m.groups()))
        except ValueError:
            return None
    else:
        t = re.sub(r"\b(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*\.?",
                   lambda x: ES_A_EN[x.group(1).lower()], texto.lower().replace(" de ", " "))
        try:
            f = dparser.parse(t, dayfirst=True, default=datetime(ref.year, 1, 1))
        except (ValueError, OverflowError):
            return None
        if not re.search(r"\d{4}|\d{1,2}[/.-]\d{1,2}[/.-]\d{2}\b", texto):
            candidatas = []
            for anio in (ref.year - 1, ref.year, ref.year + 1):
                try:
                    candidatas.append(f.replace(year=anio))
                except ValueError:                       # 29/02 en año no bisiesto
                    continue
            f = min(candidatas, key=lambda x: abs(x - ref)) if candidatas else f
    return f if abs((f - ref).days) <= MAX_DESFASE_DIAS else None


def mensajes_del_hilo(cuerpo: str) -> list:
    """El cuerpo partido en mensajes, del más nuevo (arriba) al más viejo."""
    lineas = str(cuerpo).splitlines()

    def corta(i: int, linea: str) -> bool:
        if i == 0:
            return False
        if RE_SEPARADOR.match(linea):
            return True
        if RE_CAB_DE.match(linea) and any(RE_CAB_ENVIO.match(x) for x in lineas[i + 1:i + 4]):
            return True
        return bool(RE_CITA_MAYOR.match(linea)) and not RE_CITA_MAYOR.match(lineas[i - 1])

    bordes = [0, *(i for i, linea in enumerate(lineas) if corta(i, linea)), len(lineas)]
    return ["\n".join(lineas[a:b]) for a, b in zip(bordes, bordes[1:]) if b > a]


def texto_propio(asunto: str, cuerpo: str) -> str:
    """Asunto + lo que escribió quien manda el correo (sin el hilo citado)."""
    mensajes = mensajes_del_hilo(cuerpo)
    return f"{asunto}\n{mensajes[0] if mensajes else ''}"


def campos_crudos(texto: str) -> list:
    """[(campo, fecha cruda)] en orden de aparición, de arriba hacia abajo."""
    salida = []
    for linea in str(texto).splitlines():
        resto = linea
        for campo, patron in RE_CAMPO.items():
            for m in list(patron.finditer(resto)):
                salida.append((campo, m.group(1)))
                resto = resto[:m.start()] + " " * (m.end() - m.start()) + resto[m.end():]
    return salida


def _fechas_por_campo(bloques: list, ref: datetime, respuesta: bool):
    """{campo: (fecha, fuente)} del texto propio y del citado. Propio: gana el bloque de mayor rango y, dentro de
    él, lo primero que aparece. Citado: lo primero (el mensaje citado más nuevo)."""
    rango = RANGO_BLOQUE[respuesta]
    propio, citado = {}, {}
    for orden, (bloque, texto) in enumerate(bloques):
        es_propio = bloque != "citado"
        if es_propio and bloque not in rango:            # asunto de una respuesta
            continue
        destino = propio if es_propio else citado
        for campo, crudo in campos_crudos(texto):
            fecha = a_fecha(crudo, ref)
            if fecha is None:
                continue
            clave_orden = (-rango[bloque], orden) if es_propio else (0, orden)
            if campo not in destino or clave_orden < destino[campo][2]:
                destino[campo] = (fecha, f"{bloque} · {campo} · {crudo}", clave_orden)
    return propio, citado


def leer_correo(tipo: str, asunto: str, cuerpo: str, pdfs: list, fecha_mail: datetime, respuesta: bool):
    """Lo que dice UN correo: ETA, ATA y ATD con su fuente, recepción 3PL, puerto, POL y buque. None si no trae nada."""
    datos = {"tipo": tipo, "respuesta": bool(respuesta)}
    if tipo == "recepcion_dc":
        m = RE_FECHA.search(str(asunto))
        datos["recepcion_dc"] = (a_fecha(m.group(0), fecha_mail) if m else None) or fecha_mail
        return datos
    bloques = ([("asunto", asunto)]
               + [("cuerpo" if i == 0 else "citado", t) for i, t in enumerate(mensajes_del_hilo(cuerpo))]
               + [("pdf", t) for t in pdfs])
    propio, citado = _fechas_por_campo(bloques, fecha_mail, respuesta)
    dia = pd.Timestamp(fecha_mail).normalize()

    def toma(tabla, campo, cuando=lambda f: True):
        return tabla[campo] if campo in tabla and cuando(pd.Timestamp(tabla[campo][0])) else None

    futura, no_futura = (lambda f: f > dia), (lambda f: f <= dia)
    elegidas = {"eta": toma(propio, "new_eta") or toma(propio, "eta_ata", futura) or toma(propio, "eta"),
                "ata": toma(propio, "ata") or toma(propio, "eta_ata", no_futura),
                "atd": (toma(propio, "atd") or toma(propio, "etd_atd", no_futura)
                        or toma(citado, "atd") or toma(citado, "etd_atd", no_futura))}
    for campo, valor in elegidas.items():
        datos[campo], datos[f"{campo}_fuente"] = (valor[0], valor[1]) if valor else (None, None)
    puerto, locode = RE_PUERTO.search("\n".join([*pdfs, str(cuerpo)])), RE_LOCODE_PE.search(str(asunto))
    datos["puerto"] = (re.sub(r"\s+", " ", puerto.group(1)).strip().upper() if puerto
                       else LOCODE_PE[locode.group(1)] if locode else None)
    for campo, patron in (("pol", RE_POL), ("buque", RE_BUQUE)):
        m = patron.search(str(asunto))
        datos[campo] = m.group(1).strip().upper() if m else None
    sin_nada = all(datos.get(c) is None for c in ("eta", "ata", "atd", "puerto", "pol", "buque"))
    return None if sin_nada else datos


def _ultimo(df: pd.DataFrame, col: str, extras=()) -> pd.DataFrame:
    """Fila más reciente por BL que trae `col` (df ya ordenado por fecha_mail)."""
    return (df.dropna(subset=[col]).drop_duplicates("hawb_key", keep="last")
              .set_index("hawb_key")[[col, *extras]])


def resolver(historial: pd.DataFrame) -> pd.DataFrame:
    """Una fila por hawb_key con ETA vigente, ATA, ATD, recepción 3PL, puerto, POL y buque, y la fuente de cada fecha."""
    if historial.empty:
        return pd.DataFrame(columns=COLS_RESUELTO).rename_axis("hawb_key")
    faltan = [c for c in ("hawb_key", "fecha_mail", "tipo") if c not in historial.columns]
    assert not faltan, f"correos_fwd.resolver: el historial no trae {faltan}"
    opcionales = ["respuesta", "eta", "eta_fuente", "ata", "ata_fuente", "atd", "atd_fuente", "recepcion_dc",
                  "puerto", "pol", "buque"]
    h = historial.reindex(columns=list(dict.fromkeys([*historial.columns, *opcionales])))
    for col in ("fecha_mail", "eta", "ata", "atd", "recepcion_dc"):      # sin ningun ATA la columna llega como object
        h[col] = pd.to_datetime(h[col], errors="coerce")
    h = h.sort_values("fecha_mail", kind="stable")
    con_eta = h.dropna(subset=["eta"])
    principal = con_eta[con_eta["tipo"].isin(TIPOS_PRINCIPALES) | con_eta["respuesta"].fillna(False).astype(bool)]
    vigente = (pd.concat([principal, con_eta[~con_eta["hawb_key"].isin(principal["hawb_key"])]])
                 .sort_values("fecha_mail", kind="stable"))
    eta = _ultimo(vigente, "eta", ("eta_fuente", "tipo")).rename(columns={"tipo": "eta_tipo"})

    ata = _ultimo(h, "ata", ("ata_fuente", "fecha_mail"))
    etas_despues = con_eta[["hawb_key", "fecha_mail", "eta"]].join(ata, on="hawb_key", how="inner", rsuffix="_ata")
    anula = etas_despues[(etas_despues["fecha_mail"] >= etas_despues["fecha_mail_ata"])
                         & (etas_despues["eta"] > etas_despues["ata"] + timedelta(days=DIAS_ETA_ANULA_ATA))]
    ata["ata_anulada"] = ata.index.isin(anula["hawb_key"])
    ata.loc[ata["ata_anulada"], "ata"] = pd.NaT
    ata = ata.drop(columns="fecha_mail")

    salida = eta.join(ata, how="outer").join(_ultimo(h, "atd", ("atd_fuente",)), how="outer")
    salida = salida.join(h.groupby("hawb_key")["recepcion_dc"].min(), how="outer")
    for col in ("puerto", "pol", "buque"):
        salida = salida.join(_ultimo(h, col), how="outer")
    return salida.reindex(columns=COLS_RESUELTO).rename_axis("hawb_key")


# ══════════════════════════ LECTURA DE BI E HISTORIAL ════════════════════════
def leer_bi():
    """Frescura del bi = hora en que job_in_transit lo escribio. El mtime solo sirve mientras la vista
    no existe: este job tambien escribe el archivo, y despues de su primera corrida el mtime es suyo.
    Por eso la hora real del bi viaja en la columna corte_bi de la vista."""
    assert C.RUTA_IN_TRANSIT.exists(), f"No existe {C.RUTA_IN_TRANSIT} — correr job_in_transit primero"
    mtime = C.RUTA_IN_TRANSIT.stat().st_mtime
    with pd.ExcelFile(C.RUTA_IN_TRANSIT) as libro:
        assert C.HOJA_BI in libro.sheet_names, f"{C.RUTA_IN_TRANSIT.name} sin hoja {C.HOJA_BI!r}: {libro.sheet_names}"
        bi = pd.read_excel(libro, C.HOJA_BI, dtype={c: str for c in COLS_ID_BI})
        hu = (pd.read_excel(libro, C.HOJA_HU, dtype={"delivery_number": str, "handling_unit_id": str,
                                                     COL_HAWB: str, "delivery_deleted": str})
              if C.HOJA_HU in libro.sheet_names else None)
        if C.HOJA_VISTA in libro.sheet_names:          # este job ya corrio sobre este bi
            previa = pd.read_excel(libro, C.HOJA_VISTA, nrows=1)
            assert COL_CORTE_BI in previa.columns and len(previa), (
                f"La hoja {C.HOJA_VISTA!r} no trae {COL_CORTE_BI} — correr job_in_transit para regenerar el bi")
            corte_bi = pd.Timestamp(previa[COL_CORTE_BI].iloc[0])
        else:                                          # recien escrito por job_in_transit
            corte_bi = pd.Timestamp(datetime.fromtimestamp(mtime)).floor("s")
    edad_h = C.edad_horas(corte_bi)
    if edad_h > C.MAX_HORAS:
        raise RuntimeError(f"El bi es de {corte_bi:%Y-%m-%d %H:%M} ({edad_h:.0f} h, max {C.MAX_HORAS}) — "
                           "correr job_in_transit primero; con un bi viejo la vista mentiria")
    assert len(bi), f"La hoja {C.HOJA_BI!r} esta vacia"
    faltan = [c for c in (COL_KEY, COL_HAWB, COL_SO, COL_MOT) if c not in bi.columns]
    assert not faltan, f"La hoja {C.HOJA_BI!r} ya no trae {faltan} — cambio job_in_transit"
    return bi, hu, mtime, corte_bi


def leer_historial():
    """El historial es lo unico irrecuperable del flujo: si existe y no se puede leer, se aborta.
    Nunca se sigue con vacio (eso reescribiria el archivo solo con lo nuevo).
    Devuelve arribos, salidas, notificaciones y si la hoja notificaciones aun no existe (se respalda antes)."""
    if not C.RUTA_EVENTOS.exists():
        log.info("primera corrida: %s no existe, el historial arranca vacio", C.RUTA_EVENTOS.name)
        return (pd.DataFrame(columns=COLS_ARRIBOS), pd.DataFrame(columns=COLS_SALIDAS),
                pd.DataFrame(columns=COLS_NOTIF), False)
    try:
        with pd.ExcelFile(C.RUTA_EVENTOS) as libro:
            hojas = {nombre: pd.read_excel(libro, nombre, dtype={c: str for c in cols if c in COLS_TEXTO_HIST})
                     for nombre, cols in (("arribos", COLS_ARRIBOS), ("salidas_dc", COLS_SALIDAS))}
            notif_nueva = C.HOJA_NOTIF not in libro.sheet_names
            notif = (pd.DataFrame(columns=COLS_NOTIF) if notif_nueva else
                     pd.read_excel(libro, C.HOJA_NOTIF, dtype={c: str for c in COLS_NOTIF if c in COLS_TEXTO_HIST}))
    except Exception as e:
        raise RuntimeError(f"{C.RUTA_EVENTOS.name} existe pero no se pudo leer ({type(e).__name__}: {e}). "
                           "NO se reescribe: cerrarlo si esta abierto, o restaurar una version "
                           "desde el historial de versiones de OneDrive") from e
    for (nombre, df), cols in zip(hojas.items(), (COLS_ARRIBOS, COLS_SALIDAS)):
        agregables = NUEVAS_ARRIBOS if nombre == "arribos" else set()
        faltan = [c for c in cols if c not in df.columns]
        nuevas = [c for c in faltan if c in agregables]
        for c in nuevas:                                 # version anterior del historial: se completa vacia
            df[c] = np.nan
        if nuevas:
            log.info("historial %s: columnas nuevas %s agregadas vacias (avisos leidos antes de esta version)",
                     nombre, nuevas)
        faltan = [c for c in faltan if c not in agregables]
        assert not faltan, f"{C.RUTA_EVENTOS.name}/{nombre} perdio columnas {faltan} — alguien lo edito a mano"
        for c in agregables:                             # numericos en ambos lados del concat: mismo tipo
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
        for c in COLS_FECHA_HIST & set(df.columns):
            df[c] = pd.to_datetime(df[c], errors="coerce")
    faltan = [c for c in COLS_NOTIF if c not in notif.columns]
    assert not faltan, f"{C.RUTA_EVENTOS.name}/{C.HOJA_NOTIF} perdio columnas {faltan} — alguien lo edito a mano"
    for c in COLS_FECHA_HIST & set(notif.columns):
        notif[c] = pd.to_datetime(notif[c], errors="coerce")
    if notif_nueva:
        log.info("historial sin hoja %s: se crea en esta corrida (antes se respalda el archivo)", C.HOJA_NOTIF)
    return hojas["arribos"][COLS_ARRIBOS], hojas["salidas_dc"][COLS_SALIDAS], notif[COLS_NOTIF], notif_nueva


def con_historial(bi: pd.DataFrame) -> pd.DataFrame:
    """bi + las lineas que Transito borro en los ultimos DIAS_VENTANA_SALIDOS dias. Un aereo sale del reporte al llegar al
    terminal, justo cuando llega su aviso: sin esto el aviso no encuentra su HAWB. Solo sirve para cruzar correos;
    la vista (seguimiento) sigue siendo el bi de hoy."""
    cols = [COL_HAWB, COL_SO, COL_MOT]
    if not C.RUTA_HISTORIAL.exists():
        log.warning("sin %s: los avisos solo se cruzan con el bi de hoy (correr job_in_transit actualizado)",
                    C.RUTA_HISTORIAL.name)
        return bi[cols]
    try:
        hist = pd.read_parquet(C.RUTA_HISTORIAL, columns=[COL_KEY, *cols, "visto_hasta"])
    except Exception as e:                       # paso 4 (A2): regla 14, lectura estricta. Seguir movia la ventana
        raise RuntimeError(f"{C.RUTA_HISTORIAL.name} ilegible ({type(e).__name__}: {e}) — no se cruzan avisos "
                           "solo con el bi de hoy: revisar el archivo (OneDrive) y volver a correr") from e
    desde = pd.Timestamp.now() - pd.Timedelta(days=C.DIAS_VENTANA_SALIDOS)
    fuera = hist[~hist[COL_KEY].astype(str).isin(set(bi[COL_KEY].astype(str)))
                 & pd.to_datetime(hist["visto_hasta"], errors="coerce").ge(desde)]
    salidos = set(C.clave_hawb(fuera[COL_HAWB]).dropna()) - set(C.clave_hawb(bi[COL_HAWB]).dropna())
    log.info("historial: +%d HAWB/BL que salieron del bi en los ultimos %d dias, para cruzar avisos y correos FWD",
             len(salidos), C.DIAS_VENTANA_SALIDOS)
    return pd.concat([bi[cols], fuera[cols]], ignore_index=True)


def indice_hawb(bi: pd.DataFrame):
    """hawb_key -> texto original, y hawb_key -> todas sus SO y MOT (un BL consolidado lleva varias SO)."""
    con = bi.assign(_k=C.clave_hawb(bi[COL_HAWB])).dropna(subset=["_k"])
    unir = lambda s: ";".join(sorted(set(s.dropna().astype(str).str.replace(r"\.0+$", "", regex=True).str.strip())))
    sello = con.groupby("_k").agg(so=(COL_SO, unir), mot=(COL_MOT, unir))   # lambda en agg: decenas de HAWB
    texto = con.drop_duplicates("_k").set_index("_k")[COL_HAWB].astype(str).str.strip()
    return texto.to_dict(), sello.to_dict("index")


def inicio_ventana(arribos: pd.DataFrame, salidas: pd.DataFrame):
    """Anclada en el ultimo mail registrado, no en hoy: si el job no corre varios dias no se pierde nada."""
    fechas = pd.concat([arribos["fecha_mail"], salidas["fecha_mail"]]).dropna()
    if fechas.empty:
        return datetime.now() - timedelta(days=DIAS_VENTANA_INICIAL), f"inicial de {DIAS_VENTANA_INICIAL} dias"
    return (fechas.max() - pd.Timedelta(days=MARGEN_DIAS)).to_pydatetime(), "anclada en el ultimo mail registrado"


def inicio_ventana_fwd(notif: pd.DataFrame):
    """Los correos FWD tienen su propia ancla: la primera vez leen DIAS_VENTANA_FWD (prealertas de embarques viejos)."""
    fechas = notif["fecha_mail"].dropna()
    if fechas.empty:
        return datetime.now() - timedelta(days=DIAS_VENTANA_FWD), f"inicial de {DIAS_VENTANA_FWD} dias"
    return (fechas.max() - pd.Timedelta(days=MARGEN_DIAS)).to_pydatetime(), "anclada en el ultimo correo FWD registrado"


# ══════════════════════════ OUTLOOK ══════════════════════════════════════════
def limpiar_asunto(asunto) -> str:
    sin_tags = re.sub(r"\[.*?\]", "", str(asunto))
    return re.sub(r"\b(RE|FW|RV|R|FWD)\s*:", "", sin_tags, flags=re.I).strip().upper()


def id_mensaje(item) -> str:
    """Internet Message-ID: estable aunque el mail se mueva o se copie entre Inbox y Shipments."""
    try:
        valor = item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        if valor:
            return str(valor).strip()
    except Exception:                                   # algunos items internos no lo traen
        pass
    return str(item.EntryID)


def escanear_outlook(desde: datetime, pendientes: set, ids_programacion_vistos: set,
                     desde_fwd: datetime, claves_mar: set, claves_todas: set, ids_fwd_vistos: set,
                     salidos: set = frozenset(), desde_salidos: datetime = None, intentados: set = frozenset()):
    """Una pasada: clasifica por asunto. Los adjuntos se abren despues, solo de lo elegido.
    Correos FWD (tambien respuestas: ahi llegan las ETA nuevas) y recepcion 3PL: desde su propia ventana, si nombran
    en asunto o cuerpo un BL maritimo del bi (3PL: cualquier HAWB/BL del bi). Se anota que BL aparecen solo en el
    texto citado, y que asuntos nombran un BL maritimo sin ser de un tipo conocido (alarma de formato nuevo)."""
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    desde_salidos = desde_salidos or desde
    filtro = f"[ReceivedTime] >= '{min(desde, desde_fwd, desde_salidos).strftime('%m/%d/%Y %I:%M %p')}'"
    avisos, programaciones, correos_fwd, sin_tipo, omitidos = {}, [], [], {}, 0
    for carpeta in C.carpetas_outlook(ns):
        try:
            items = carpeta.Items.Restrict(filtro)
            items.Sort("[ReceivedTime]", False)
        except Exception as e:
            log.error("No se pudo filtrar %s: %s", getattr(carpeta, "Name", "?"), e)
            continue
        for item in items:
            try:
                if getattr(item, "Class", 0) != 43:     # 43 = MailItem
                    continue
                asunto = getattr(item, "Subject", "") or ""
                recibido = item.ReceivedTime.replace(tzinfo=None)
                tipo_fwd = tipo_de(asunto)
                if tipo_fwd and recibido >= desde_fwd:
                    id_mail = id_mensaje(item)
                    if id_mail not in ids_fwd_vistos:              # ya registrado, o visto en otra carpeta
                        ids_fwd_vistos.add(id_mail)
                        cuerpo = getattr(item, "Body", "") or ""
                        plano = re.sub(r"[^a-z0-9]", "", f"{asunto} {cuerpo}".lower())
                        hits = [k for k in (claves_todas if tipo_fwd == "recepcion_dc" else claves_mar) if k in plano]
                        if hits:
                            propio = re.sub(r"[^a-z0-9]", "", texto_propio(asunto, cuerpo).lower())
                            correos_fwd.append({"mail": item, "tipo": tipo_fwd, "fecha_mail": recibido, "id_mail": id_mail,
                                               "asunto": asunto, "cuerpo": cuerpo, "respuesta": C.es_respuesta(asunto),
                                               "hits": hits, "solo_citado": {k for k in hits if k not in propio}})
                elif (recibido >= desde_fwd and not C.ASUNTO_ARRIBO.search(asunto)
                      and ASUNTO_ENVIO not in limpiar_asunto(asunto)):
                    en_asunto = re.sub(r"[^a-z0-9]", "", asunto.lower())
                    if any(k in en_asunto for k in claves_mar):
                        sin_tipo[id_mensaje(item)] = asunto
                if C.es_respuesta(asunto):
                    continue
                if C.ASUNTO_ARRIBO.search(asunto):
                    for clave in C.ids_en_asunto(asunto) & pendientes:
                        if recibido < (desde_salidos if clave in salidos else desde):
                            continue            # lo que ya salio del bi se busca mas atras: su aviso pudo llegar antes
                        if (clave, id_mensaje(item)) in intentados:
                            continue            # paso 4 (A5): este mismo correo ya se leyo y su PDF no tenia fecha
                        if clave not in avisos or recibido > avisos[clave]["fecha_mail"]:
                            avisos[clave] = {"mail": item, "fecha_mail": recibido, "id_mail": id_mensaje(item)}
                elif recibido >= desde and ASUNTO_ENVIO in limpiar_asunto(asunto):
                    id_mail = id_mensaje(item)
                    if id_mail not in ids_programacion_vistos:
                        ids_programacion_vistos.add(id_mail)
                        programaciones.append({"mail": item, "fecha_mail": recibido,
                                               "id_mail": id_mail, "asunto": asunto})
            except Exception as e:
                omitidos += 1
                log.warning("item omitido (%s)", type(e).__name__)
    log.info("Outlook: %d avisos de HAWB pendientes · %d programaciones nuevas · %d correos FWD/3PL nuevos · "
             "%d items omitidos", len(avisos), len(programaciones), len(correos_fwd), omitidos)
    if sin_tipo:
        log.info("   %d correos nombran un BL maritimo en el asunto sin ser prealerta, notificacion, aviso ni 3PL "
                 "(si alguno es de FWD, cambio su asunto): %s", len(sin_tipo),
                 " | ".join(a[:70] for a in list(sin_tipo.values())[:5]))
    return avisos, programaciones, correos_fwd


# ══════════════════════════ AVISO DE LLEGADA (PDF) ═══════════════════════════
def textos_pdf(mail) -> list:
    """Texto de cada PDF adjunto, por separado: cada uno puede venir de otro sistema."""
    textos = []
    try:
        for adjunto in mail.Attachments:
            if not str(adjunto.FileName).lower().endswith(".pdf"):
                continue
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                ruta_tmp = tmp.name
            try:
                adjunto.SaveAsFile(ruta_tmp)
                with pdfplumber.open(ruta_tmp) as pdf:
                    textos.append("\n".join(p.extract_text() or "" for p in pdf.pages))
            except Exception as e:
                log.warning("   PDF %r ilegible: %s", adjunto.FileName, type(e).__name__)
            finally:
                if os.path.exists(ruta_tmp):
                    os.remove(ruta_tmp)
    except Exception as e:
        log.warning("   adjuntos ilegibles: %s", type(e).__name__)
    return textos


def texto_pdfs(mail) -> str:
    return "\n".join(textos_pdf(mail)).strip()


def fecha_llegada(texto: str):
    for patron in PATRONES_LLEGADA:
        m = patron.search(texto)
        if m:
            return pd.to_datetime(m.group(1), format="%d/%m/%Y", errors="coerce")
    return pd.NaT


def extraer_bultos(texto: str) -> dict:
    """Logica validada de rpa_maestro. Fila 'Bultos a b c d e' -> total=a, buenos=c, malos=d, faltantes=e
    (el segundo numero no se mapea: el original tampoco). Respaldo: 'Manifestados .. Buenos .. Malos ..'.
    Si el PDF no trae ninguna de las dos formas, quedan vacios: no se inventa un cero."""
    res = dict.fromkeys(COLS_BULTOS)
    m = RE_BULTOS_FILA.search(texto)
    if m:
        n = [int(x) for x in m.groups()]
        res.update(bultos_total=n[0], bultos_buenos=n[2], bultos_malos=n[3], bultos_faltantes=n[4])
        return res
    m = RE_BULTOS_ETIQUETAS.search(texto)
    if m:
        n = [int(x) for x in m.groups()]
        res.update(bultos_total=n[0], bultos_buenos=n[1], bultos_malos=n[2])
    return res


def procesar_avisos(avisos: dict, texto_hawb: dict, sello: dict, ahora) -> pd.DataFrame:
    filas = []
    for clave, info in avisos.items():
        texto = texto_pdfs(info["mail"])
        if not texto:                    # escaneado o sin PDF: el aviso igual dice que llego (llegada = su correo)
            log.warning("   %s: aviso sin PDF legible — se registra igual, con la fecha del correo como llegada y "
                        "sin bultos; el PDF se reintenta en la proxima corrida", texto_hawb[clave])
        llegada = fecha_llegada(texto)
        if texto and pd.isna(llegada):
            con_fecha = [re.sub(r"\s+", " ", l).strip() for l in texto.splitlines() if RE_FECHA_SUELTA.search(l)][:4]
            log.warning("   %s: PDF sin fecha de llegada reconocible (se reintenta en la proxima corrida) — "
                        "lineas del PDF con fecha: %s", texto_hawb[clave], con_fecha or "ninguna")
        filas.append({"hawb": texto_hawb[clave], "hawb_key": clave, "id_mail": info["id_mail"],
                      "fecha_mail": info["fecha_mail"], "fecha_llegada": llegada,
                      **extraer_bultos(texto),
                      **sello.get(clave, {}), "capturado": ahora})
    log.info("   avisos registrados: %d de %d", len(filas), len(avisos))
    nuevos = pd.DataFrame(filas, columns=COLS_ARRIBOS)
    nuevos[COLS_BULTOS] = nuevos[COLS_BULTOS].astype("float64")      # aviso sin bultos: NaN, no None
    return nuevos


# ══════════════════════════ CORREOS FWD Y RECEPCION 3PL ═══════════════════════
def procesar_correos_fwd(correos: list, texto_hawb: dict, sello: dict, ahora) -> pd.DataFrame:
    """Un registro por HAWB/BL × correo: lo que dice el correo y de donde salio cada fecha (leer_correo).
    - BL que solo aparece en el texto citado: de ese correo solo vale el ATD (la ETA es de otro embarque o vieja).
    - Correo que nombra varios BL: sus fechas valen para todos y la fuente lo dice ('correo de N BL').
    - Correo sin ninguna fecha ni dato legible: no se registra y se avisa (suele ser un formato nuevo)."""
    filas, sin_datos, solo_citado, varios = [], [], 0, 0
    for c in correos:
        pdfs = textos_pdf(c["mail"]) if c["tipo"] in TIPOS_CON_PDF else []
        datos = leer_correo(c["tipo"], c["asunto"], c["cuerpo"], pdfs, c["fecha_mail"], c["respuesta"])
        if datos is None:
            sin_datos.append(c["asunto"])
            continue
        n_bl = len(set(c["hits"]) - c["solo_citado"]) if c["tipo"] != "recepcion_dc" else 1
        varios += n_bl > 1
        for k in c["hits"]:
            fila = {"hawb": texto_hawb[k], "hawb_key": k, "id_mail": c["id_mail"], "fecha_mail": c["fecha_mail"],
                    **datos, **sello.get(k, {}), "capturado": ahora}
            if k in c["solo_citado"] and c["tipo"] != "recepcion_dc":
                fila.update(eta=None, eta_fuente=None, ata=None, ata_fuente=None)
                solo_citado += 1
                if fila.get("atd") is None:
                    continue
            elif n_bl > 1:
                for campo in ("eta_fuente", "ata_fuente", "atd_fuente"):
                    if fila.get(campo):
                        fila[campo] += f" · correo de {n_bl} BL"
            filas.append(fila)
    nuevos = pd.DataFrame(filas, columns=COLS_NOTIF)
    for col in COLS_FECHA_HIST & set(COLS_NOTIF):        # tipadas aunque ningun correo traiga ATA
        nuevos[col] = pd.to_datetime(nuevos[col], errors="coerce")
    log.info("   correos FWD/3PL: %d registros (%s) · %d BL solo en texto citado (sin ETA) · %d correos de varios BL",
             len(nuevos), nuevos["tipo"].value_counts().to_dict(), solo_citado, varios)
    if sin_datos:
        log.warning("   %d correos FWD/3PL nombran un BL del bi pero no traen ninguna fecha legible (¿formato nuevo?): %s",
                    len(sin_datos), " | ".join(a[:70] for a in sin_datos[:5]))
    return nuevos


# ══════════════════════════ PROGRAMACION FWD (tabla del mail) ═════════════════
def _canonizar_header(crudos):
    resultado = []
    tiene_bl = any(_ALIAS_LOOKUP.get(str(h).upper().strip().rstrip(":")) == "BL" for h in crudos)
    for i, h in enumerate(crudos):
        limpio = str(h).upper().strip().rstrip(":")
        if not limpio:
            resultado.append(f"_BLANK_{i}")
            continue
        canon = _ALIAS_LOOKUP.get(limpio, limpio)
        resultado.append("BL" if canon == "AWB" and not tiene_bl else canon)
    return resultado


def _es_tabla_envios(encabezado) -> bool:
    return len({"ORDEN", "AWB", "BL", "FACTURA", "CLIENTE", "TERMINAL", "DESTINO"} & set(encabezado)) >= 2


def _celda_html(td) -> str:
    texto = td.get_text(separator="\n").replace("\xa0", " ")
    return " ".join(t.strip() for t in texto.split("\n") if t.strip())


def extraer_tabla_html(html):
    if not html:
        return None
    try:
        sopa = BeautifulSoup(html, "lxml")
    except Exception:                                   # sin lxml: parser estandar
        sopa = BeautifulSoup(html, "html.parser")
    for tabla in sopa.find_all("table"):
        if tabla.find("table"):                         # tabla de layout que envuelve otras
            continue
        filas = tabla.find_all("tr")
        if len(filas) < 2:
            continue
        celdas = filas[0].find_all(["th", "td"])
        if not celdas:
            continue
        encabezado = _canonizar_header([_celda_html(c) for c in celdas])
        if not _es_tabla_envios(encabezado):
            continue
        registros = []
        for fila in filas[1:]:
            valores = [_celda_html(td) for td in fila.find_all("td")]
            if not valores:
                continue
            valores += [""] * (len(encabezado) - len(valores))
            registro = dict(zip(encabezado, valores[:len(encabezado)]))
            if any(v not in ("", "nan") for v in registro.values()):
                registros.append(registro)
        if registros:
            return pd.DataFrame(registros)
    return None


def extraer_tabla_texto(cuerpo):
    if not cuerpo:
        return None
    lineas = [l.strip() for l in cuerpo.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    lineas = [l for l in lineas if l and not set(l).issubset(set("\t \xa0"))]
    inicio = None
    for i, linea in enumerate(lineas):
        if linea.upper().rstrip(":") in HEADERS_VALIDOS:
            siguientes = {lineas[i + j].upper().rstrip(":") for j in range(min(6, len(lineas) - i))}
            if len(HEADERS_VALIDOS & siguientes) >= 3:
                inicio = i
                break
    if inicio is None:
        return None
    crudos, i = [], inicio
    while i < len(lineas) and lineas[i].upper().rstrip(":") in HEADERS_VALIDOS:
        crudos.append(lineas[i])
        i += 1
    encabezado = _canonizar_header(crudos)
    if not _es_tabla_envios(encabezado):
        return None
    valores = []
    while i < len(lineas) and not any(sw in lineas[i].upper() for sw in STOPWORDS_CORTE):
        valores.append(lineas[i])
        i += 1
    if not valores:
        return None
    pos_factura = encabezado.index("FACTURA") if "FACTURA" in encabezado else None
    exige_orden = encabezado[0] == "ORDEN"
    registros, j = [], 0
    while j < len(valores):
        if exige_orden and not ORDEN_PATTERN.match(valores[j]):
            break
        fila = []
        for col in range(len(encabezado)):
            if j >= len(valores):
                fila.append("")
            elif col == pos_factura:                    # una fila puede traer varias facturas seguidas
                facturas = []
                while j < len(valores) and RE_FACTURA.match(valores[j]):
                    facturas.append(valores[j])
                    j += 1
                fila.append(" ".join(facturas))
            else:
                fila.append(valores[j])
                j += 1
        if any(fila):
            registros.append(dict(zip(encabezado, fila)))
    return pd.DataFrame(registros) if registros else None


def hawbs_de_tabla(tabla: pd.DataFrame) -> dict:
    """BL si es valido, si no AWB (regla del flujo validado). Una llave por HAWB por mail."""
    vacio = pd.Series("", index=tabla.index)
    encontrados = {}
    for v_bl, v_awb in zip(tabla.get("BL", vacio).astype(str), tabla.get("AWB", vacio).astype(str)):
        for valor in (v_bl, v_awb):
            clave = C.normalizar_id(valor)
            if len(clave) >= C.LARGO_MIN_ID:
                encontrados.setdefault(clave, valor.strip())
                break
    return encontrados


def fecha_programada(dia_mes, recibido):
    """'dd/mm' del cuerpo sin año: se toma el año que deja la fecha mas cerca del mail.
    Cubre los dos cruces de año (mail del 30/12 para el 02/01 y mail del 02/01 que cita el 30/12)."""
    if not dia_mes:
        return pd.NaT
    dia, mes = (int(x) for x in re.split(r"[/-]", dia_mes))
    base = pd.Timestamp(recibido)
    candidatas = []
    for anio in (base.year - 1, base.year, base.year + 1):
        try:
            candidatas.append(pd.Timestamp(year=anio, month=mes, day=dia))
        except ValueError:                              # 31/02, 29/02 en año no bisiesto: no se inventa
            continue
    return min(candidatas, key=lambda f: abs(f - base)) if candidatas else pd.NaT


def procesar_programaciones(programaciones: list, sello: dict, ahora) -> pd.DataFrame:
    filas = []
    for p in programaciones:
        item = p["mail"]
        cuerpo = getattr(item, "Body", "") or ""
        tabla = extraer_tabla_html(getattr(item, "HTMLBody", "") or "")
        if tabla is None:
            tabla = extraer_tabla_texto(cuerpo)
        hawbs = hawbs_de_tabla(tabla) if tabla is not None and not tabla.empty else {}
        if not hawbs:
            log.warning("   programacion sin HAWB/BL legible: %s", p["asunto"][:60])
            continue
        plano = cuerpo.replace("\r", " ").replace("\n", " ")
        m_fecha, m_hora = RE_FECHA_PROG.search(plano), RE_HORA_PROG.search(plano)
        fecha_prog = fecha_programada(m_fecha.group(1) if m_fecha else None, p["fecha_mail"])
        hora = m_hora.group(1) if m_hora else None
        for clave, original in hawbs.items():
            filas.append({"hawb": original, "hawb_key": clave, "id_mail": p["id_mail"],
                          "fecha_mail": p["fecha_mail"], "fecha_programada": fecha_prog,
                          "hora_estimada": hora, **sello.get(clave, {}), "capturado": ahora})
        log.info("   %s: %d HAWB/BL", p["asunto"][:50], len(hawbs))
    return pd.DataFrame(filas, columns=COLS_SALIDAS)


# ══════════════════════════ HISTORIAL Y VISTA ════════════════════════════════
def anexar(historial: pd.DataFrame, nuevas: pd.DataFrame, columnas: list):
    """Append idempotente por hawb_key + id_mail: correr dos veces no duplica."""
    if nuevas.empty:
        return historial, 0
    llave = ["hawb_key", "id_mail"]
    nuevas = nuevas.drop_duplicates(llave)
    existentes = pd.MultiIndex.from_frame(historial[llave].astype(str))
    ya = pd.MultiIndex.from_frame(nuevas[llave].astype(str)).isin(existentes)
    nuevas = nuevas.loc[~ya]
    if nuevas.empty:
        return historial, 0
    combinado = nuevas if historial.empty else pd.concat([historial, nuevas], ignore_index=True)
    return combinado[columnas], len(nuevas)


def escribir_historial(arribos: pd.DataFrame, salidas: pd.DataFrame, notif: pd.DataFrame) -> None:
    tmp = C.RUTA_EVENTOS.with_name(C.RUTA_EVENTOS.stem + ".tmp.xlsx")   # .xlsx al final: openpyxl lo exige
    with pd.ExcelWriter(tmp, engine="openpyxl") as w:
        arribos.to_excel(w, sheet_name="arribos", index=False)
        salidas.to_excel(w, sheet_name="salidas_dc", index=False)
        notif.to_excel(w, sheet_name=C.HOJA_NOTIF, index=False)
    C.reemplazar_atomico(tmp, C.RUTA_EVENTOS)


def respaldar_historial() -> None:
    """Copia del historial antes de su primera escritura con la hoja notificaciones (cambia la estructura)."""
    destino = C.RUTA_EVENTOS.with_name(f"{C.RUTA_EVENTOS.stem}_respaldo_{datetime.now():%Y%m%d_%H%M}.xlsx")
    shutil.copy2(C.RUTA_EVENTOS, destino)
    log.info("historial respaldado en %s", destino.name)


def construir_seguimiento(bi: pd.DataFrame, arribos: pd.DataFrame, salidas: pd.DataFrame, kn: pd.DataFrame,
                          corte_bi) -> pd.DataFrame:
    ultimo_aviso = (arribos.sort_values("fecha_mail").drop_duplicates("hawb_key", keep="last")
                    .set_index("hawb_key")[["fecha_llegada", "fecha_mail"]]
                    .set_axis(["llegada_fecha", "llegada_aviso_mail"], axis=1))
    ultima_salida = (salidas.sort_values("fecha_mail").drop_duplicates("hawb_key", keep="last")
                     .set_index("hawb_key")[["fecha_programada", "hora_estimada", "fecha_mail"]]
                     .set_axis(["salida_dc_fecha", "salida_dc_hora", "salida_dc_mail"], axis=1))

    clave = C.clave_hawb(bi[COL_HAWB])
    reprogramados = int(salidas.loc[salidas["hawb_key"].isin(clave.dropna()), "hawb_key"]
                        .value_counts().gt(1).sum())
    if reprogramados:
        log.info("   %d HAWB con mas de una programacion: se muestra la ultima", reprogramados)

    correo = kn[["atd", "eta", "ata", "puerto", "recepcion_dc"]].set_axis(
        ["atd_correo", "eta_correo", "_ata_correo", "puerto_descarga", "recepcion_dc"], axis=1)
    vista = (bi.assign(_k=clave.fillna(""))
               .join(ultimo_aviso, on="_k").join(ultima_salida, on="_k").join(correo, on="_k").drop(columns="_k"))
    assert len(vista) == len(bi), "El cruce por HAWB multiplico lineas — revisar unicidad del historial"

    # Aviso sin fecha legible en el PDF: el correo del aviso llega con la carga -> su fecha es la llegada.
    # El historial no se toca (sigue sin fecha y se reintenta); solo la vista muestra el respaldo y lo dice.
    con_aviso = vista["llegada_aviso_mail"].notna()
    del_pdf = pd.to_datetime(vista["llegada_fecha"], errors="coerce")
    vista["llegada_fecha"] = del_pdf.fillna(pd.to_datetime(vista["llegada_aviso_mail"]).dt.normalize())
    vista["llegada_fecha_fuente"] = np.select([con_aviso & del_pdf.notna(), con_aviso],
                                              ["PDF del aviso", "fecha del correo"], default=None)

    # Arribo aereo: aviso o ATA del reporte
    ata_reporte = (pd.to_datetime(vista[COL_ATA], errors="coerce") if COL_ATA in vista.columns
                   else pd.Series(pd.NaT, index=vista.index))
    arribo = con_aviso | ata_reporte.notna()

    # Arribo maritimo: ATA del correo, ATA del reporte o ETA vigente ya vencida. El aviso no cuenta: tambien llega
    # cuando la carga sale de planta, con la ETA a futuro. El ATA por correo casi no llega (0 de 16 al 17/11/2026)
    corte_dia = pd.Timestamp(corte_bi).normalize()
    maritimo = vista[COL_MOT].astype("string").str.upper().str.strip().isin(C.MOT_MARITIMO).fillna(False).astype(bool)
    ata_correo = pd.to_datetime(vista["_ata_correo"], errors="coerce")
    eta = pd.to_datetime(vista["eta_correo"], errors="coerce")
    eta_vencida = eta.where(eta.dt.normalize() <= corte_dia)
    llegada_mar = ata_correo.fillna(ata_reporte).fillna(eta_vencida)
    fuente_mar = pd.Series(np.select([ata_correo.notna(), ata_reporte.notna(), eta_vencida.notna()],
                                     ["ATA del correo", "ATA del reporte", "ETA vencida"], default=None), index=vista.index)
    vista["llegada_fecha"] = llegada_mar.where(maritimo, vista["llegada_fecha"])
    vista["llegada_fecha_fuente"] = fuente_mar.where(maritimo, vista["llegada_fecha_fuente"])
    arribo = np.where(maritimo, llegada_mar.notna(), arribo)

    # Entregado al DC: recepcion 3PL, o salida FWD hace mas de DIAS_HABILES_ENTREGA dias habiles
    con_salida = vista["salida_dc_mail"].notna()
    salida = pd.to_datetime(vista["salida_dc_fecha"], errors="coerce").fillna(
        pd.to_datetime(vista["salida_dc_mail"], errors="coerce").dt.normalize())
    entregado = vista["recepcion_dc"].notna() | (con_salida & (C.dias_habiles(salida, corte_dia) > C.DIAS_HABILES_ENTREGA))
    vista["hito"] = np.select(
        [clave.isna(), entregado, con_salida, arribo],
        ["Sin HAWB", "Entregado al DC", "En tránsito a DC", "Arribado"], default="En tránsito internacional")
    vista[COL_CORTE_BI] = corte_bi
    vista = vista.drop(columns="_ata_correo")

    frente = [c for c in COLS_FRENTE if c in vista.columns]
    resto = [c for c in vista.columns if c not in frente and c not in COLS_TRAZA]
    return vista[frente + resto + COLS_TRAZA]


def medida_bulto(hu: pd.DataFrame) -> pd.Series:
    """'120×80×110' en cm enteros; '?' donde falta una dimension."""
    d = hu[["largo_cm", "ancho_cm", "alto_cm"]].round(0).astype("Int64").astype("string").fillna("?")
    return d["largo_cm"].str.cat([d["ancho_cm"], d["alto_cm"]], sep="×")


def perfil_bultos(medidas: pd.Series) -> str:
    """'8× 120×80×110 · 2× 60×40×40 · +3 otros': la mezcla fisica que llega, en una celda."""
    conteo = medidas.value_counts()
    partes = [f"{n}× {m}" for m, n in conteo.head(PERFIL_TOP).items()]
    if len(conteo) > PERFIL_TOP:
        partes.append(f"+{int(conteo.iloc[PERFIL_TOP:].sum())} otros")
    return " · ".join(partes)


def construir_embarques(vista: pd.DataFrame, hu, arribos: pd.DataFrame) -> pd.DataFrame:
    """Una fila por HAWB/BL: donde esta, cuando sale al DC y cuanto ocupa. El peso y los bultos del aviso
    viven aqui y no en seguimiento: en la linea se repetirian por cada item del mismo embarque."""
    lineas = vista[vista[COL_HAWB].notna()].assign(_embarque=lambda d: C.norm_doc(d[COL_HAWB]))
    lineas["_rango"] = lineas["hito"].map({h: i for i, h in enumerate(C.ORDEN_HITO)})
    g = lineas.groupby("_embarque")
    emb = pd.DataFrame({
        "mot": g[COL_MOT].agg(lambda s: " / ".join(sorted(s.dropna().astype(str).unique()))),  # 1-2 valores
        "_rango": g["_rango"].min(),                     # varias lineas: manda la menos avanzada
        "llegada_fecha": g["llegada_fecha"].max(),
        "salida_dc_fecha": g["salida_dc_fecha"].max(),
        "salida_dc_hora": g["salida_dc_hora"].first(),
    })
    discrepan = int((g["_rango"].nunique() > 1).sum())
    if discrepan:
        log.info("   %d embarques con lineas en hitos distintos: se muestra el menos avanzado", discrepan)

    requeridas = COLS_HU_CRUDAS + [COL_HAWB]
    faltan = requeridas if hu is None else [c for c in requeridas if c not in hu.columns]
    if faltan:
        log.warning("   hu_detail sin %s (version anterior de job_in_transit): embarques sin volumen", faltan)
    else:
        v = C.en_unidades_estandar(hu)
        v["_embarque"] = C.norm_doc(v[COL_HAWB]).fillna(SIN_HAWB)
        v["_medida"] = medida_bulto(v)
        gv = v.groupby("_embarque")
        volumen = pd.DataFrame({
            "n_dn": gv["delivery_number"].nunique(),
            "n_bultos": gv.size(),
            # min_count=1: un embarque sin ninguna medida queda vacio, no en 0 (0 kg se lee como dato)
            "peso_bruto_kg": gv["peso_bruto_kg"].sum(min_count=1).round(1),
            "peso_cobrable_kg": gv["peso_cobrable_kg"].sum(min_count=1).round(1),
            "volumen_m3": gv["volumen_m3"].sum(min_count=1).round(3),
            "bulto_mas_grande": (v.sort_values("volumen_m3", ascending=False)
                                  .drop_duplicates("_embarque").set_index("_embarque")["_medida"]),
            "perfil_bultos": gv["_medida"].agg(perfil_bultos),
        })
        emb = emb.join(volumen, how="outer")

    # bultos segun el aviso de llegada (ultimo aviso del HAWB), al lado del conteo de Lake
    ultimo = (arribos.sort_values("fecha_mail").drop_duplicates("hawb_key", keep="last")
              .set_index("hawb_key")[COLS_BULTOS])
    clave = pd.Series(emb.index, index=emb.index).map(C.normalizar_id)
    for col in COLS_BULTOS:
        emb[col.replace("bultos_", "aviso_")] = pd.to_numeric(clave.map(ultimo[col]), errors="coerce")

    emb["hito"] = emb["_rango"].map(dict(enumerate(C.ORDEN_HITO)))
    emb.loc[emb.index == SIN_HAWB, "hito"] = "Sin HAWB"
    emb = (emb.assign(_sin=emb.index == SIN_HAWB)
              .sort_values(["_sin", "salida_dc_fecha", "_rango", "llegada_fecha"],
                           ascending=[True, True, False, True], na_position="last"))
    columnas = ["mot", "hito", "llegada_fecha", "salida_dc_fecha", "salida_dc_hora", "n_dn", "n_bultos",
                "aviso_total", "aviso_buenos", "aviso_malos", "aviso_faltantes",
                "peso_bruto_kg", "peso_cobrable_kg", "volumen_m3", "bulto_mas_grande", "perfil_bultos"]
    return emb.reindex(columns=columnas).rename_axis("embarque").reset_index()


def escribir_hojas(hojas: dict, mtime_leido: float) -> None:
    """Reemplaza solo las hojas propias; las de job_in_transit quedan intactas y las que se arman
    desde seguimiento se retiran, para que ninguna quede armada sobre una version anterior."""
    C.reemplazar_hojas_excel(
        C.RUTA_IN_TRANSIT,
        hojas,
        mtime_leido,
        C.HOJAS_AJENAS_EVENTOS,
        retirar=C.HOJAS_AGUAS_ABAJO_EVENTOS,
        logger=log,
    )


def controles_fwd(vista: pd.DataFrame) -> None:
    """Alarmas de formato en cada corrida: un ATD del correo que no cuadra con el reporte o una ETA muy lejos del
    Outlook suelen ser una plantilla nueva o una fecha mal leida. Solo avisan en el log; no cambian nada."""
    if not {"atd", "outlook_eta"} <= set(vista.columns):
        return
    mar = vista[vista[COL_MOT].astype("string").str.upper().str.strip().isin(C.MOT_MARITIMO).fillna(False).astype(bool)]
    por_bl = mar.dropna(subset=[COL_HAWB]).drop_duplicates(COL_HAWB).set_index(COL_HAWB)
    dia = lambda col: pd.to_datetime(por_bl[col], errors="coerce").dt.normalize()
    atd = (dia("atd_correo") - dia("atd")).dt.days.abs()
    eta = (dia("eta_correo") - dia("outlook_eta")).dt.days.abs()
    for nombre, dif, limite in (("ATD del correo vs ATD del reporte", atd, 1), ("ETA del correo vs Outlook", eta, 15)):
        fuera = dif[dif > limite]
        if len(fuera):
            log.warning("   %s: %d BL a mas de %d dias (¿fecha mal leida?): %s", nombre, len(fuera), limite,
                        ", ".join(f"{bl} ({int(d)} d)" for bl, d in fuera.items()))


# ══════════════════════════ DRY_RUN ══════════════════════════════════════════
def reporte_dry_run(vista: pd.DataFrame, embarques: pd.DataFrame, kn: pd.DataFrame) -> None:
    """Lo que cambiaria, para validarlo antes de escribir: hitos por embarque contra la hoja actual y, por BL
    maritimo, sus fechas con la fuente de cada una y el ATD contra el reporte."""
    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 45)
    try:
        actual = pd.read_excel(C.RUTA_IN_TRANSIT, C.HOJA_EMBARQUES, usecols=["embarque", "hito"], dtype={"embarque": str})
    except Exception:                                   # sin hoja embarques todavia: todo sale como nuevo
        actual = pd.DataFrame(columns=["embarque", "hito"])
    cruce = embarques[["embarque", "hito"]].merge(actual.assign(embarque=C.norm_doc(actual["embarque"])),
                                                   on="embarque", how="left", suffixes=("", "_hoy"))
    print("\nHitos por embarque · filas: hoja actual · columnas: esta version")
    print(pd.crosstab(cruce["hito_hoy"].fillna("(no estaba)"), cruce["hito"]).to_string())

    mar = vista[vista[COL_MOT].astype("string").str.upper().str.strip().isin(C.MOT_MARITIMO).fillna(False).astype(bool)
                & vista[COL_HAWB].notna()]
    por_bl = (mar.assign(_k=C.clave_hawb(mar[COL_HAWB])).dropna(subset=["_k"]).drop_duplicates("_k").set_index("_k")
                 .join(kn[["eta_fuente", "atd_fuente", "ata_fuente", "ata_anulada"]])
                 .sort_values("eta_correo", na_position="last"))
    dm = lambda col: pd.to_datetime(por_bl[col], errors="coerce").dt.strftime("%d/%m")
    tabla = pd.DataFrame({"hito": por_bl["hito"], "llegada": dm("llegada_fecha"),
                          "fuente_llegada": por_bl["llegada_fecha_fuente"], "eta": dm("eta_correo"),
                          "fuente_eta": por_bl["eta_fuente"], "atd_correo": dm("atd_correo"),
                          "atd_reporte": dm("atd") if "atd" in por_bl.columns else None,
                          "salida_fwd": dm("salida_dc_fecha"), "op3pl": dm("recepcion_dc")})
    tabla.index = por_bl[COL_HAWB].astype(str).str.strip()
    print("\nBL maritimos del bi: fechas con su fuente")
    print(tabla.fillna("—").to_string())
    if "atd" in por_bl.columns:
        dif = (pd.to_datetime(por_bl["atd_correo"]).dt.normalize()
               - pd.to_datetime(por_bl["atd"], errors="coerce").dt.normalize()).dt.days.abs()
        print(f"\n   ATD del correo a ±1 dia del reporte: {int((dif <= 1).sum())} de {int(dif.notna().sum())}")
    for hawb, fuente in por_bl.loc[por_bl["ata_anulada"].eq(True), [COL_HAWB, "ata_fuente"]].itertuples(index=False):
        print(f"   ATA anulada en {hawb}: {fuente} (una ETA del mismo correo o de uno posterior cae despues)")


# ══════════════════════════ MAIN ═════════════════════════════════════════════
def main_eventos() -> None:
    dry_run = C.DRY_RUN_EVENTOS
    log.info("=== EVENTOS DE TRANSITO · bi -> Outlook -> historial -> seguimiento%s ===", " · DRY_RUN" * dry_run)
    bi, hu, mtime_bi, corte_bi = leer_bi()
    log.info("bi de job_in_transit: %s", f"{corte_bi:%Y-%m-%d %H:%M}")
    arribos, salidas, notif, notif_nueva = leer_historial()

    texto_hawb, sello = indice_hawb(con_historial(bi))         # bi + lo que Transito ya borro (aereos en terminal)
    salidos_bi = set(texto_hawb) - set(C.clave_hawb(bi[COL_HAWB]).dropna())
    # un aviso guardado sin fecha de llegada cuenta como pendiente: se vuelve a leer mientras siga en la ventana
    sin_fecha = set(arribos.loc[arribos["fecha_llegada"].isna(), "hawb_key"].dropna())
    ya_con_aviso = set(arribos["hawb_key"].dropna()) - sin_fecha
    pendientes = {k for k in texto_hawb if k not in ya_con_aviso}
    log.info("bi: %d lineas · %d HAWB · %d pendientes de aviso (%d reintentos sin fecha) · %d lineas sin HAWB",
             len(bi), len(texto_hawb), len(pendientes), len(sin_fecha & pendientes),
             int(C.clave_hawb(bi[COL_HAWB]).isna().sum()))

    claves_mar = {k for k, s in sello.items() if set(str(s.get("mot", "")).upper().split(";")) & C.MOT_MARITIMO}
    desde, motivo = inicio_ventana(arribos, salidas)
    desde_fwd, motivo_fwd = inicio_ventana_fwd(notif)
    log.info("ventana Outlook desde %s (%s) · correos FWD desde %s (%s) · %d BL maritimos en el bi",
             f"{desde:%Y-%m-%d %H:%M}", motivo, f"{desde_fwd:%Y-%m-%d %H:%M}", motivo_fwd, len(claves_mar))
    avisos, programaciones, correos_fwd = escanear_outlook(
        desde, pendientes, set(salidas["id_mail"].dropna()),
        desde_fwd, claves_mar, set(texto_hawb), set(notif["id_mail"].dropna().astype(str)),
        salidos_bi, datetime.now() - timedelta(days=C.DIAS_VENTANA_SALIDOS),
        set(map(tuple, arribos.loc[arribos["fecha_llegada"].isna(), ["hawb_key", "id_mail"]].dropna().astype(str).values)))

    ahora = pd.Timestamp.now().floor("s")
    salidas_nuevas = procesar_programaciones(programaciones, sello, ahora)
    if len(salidas_nuevas):
        # llave de cruce: si casi ningun HAWB de FWD esta en el bi, bl_awb_num no es el mismo identificador
        en_bi = salidas_nuevas["hawb_key"].isin(texto_hawb.keys())
        log.info("   HAWB/BL en programaciones: %d · en el bi: %d (%.0f%%)",
                 len(salidas_nuevas), int(en_bi.sum()), 100 * en_bi.mean())
    arribos_nuevos = procesar_avisos(avisos, texto_hawb, sello, ahora)
    de_salidos = arribos_nuevos[arribos_nuevos["hawb_key"].isin(salidos_bi)]
    if len(de_salidos):
        log.info("   %d avisos de HAWB/BL que ya salieron del bi (los cruza el historial): %s", len(de_salidos),
                 ", ".join(de_salidos["hawb"].astype(str).head(15)))
    notif_nuevas = procesar_correos_fwd(correos_fwd, texto_hawb, sello, ahora)
    # el aviso que antes quedo sin fecha y ahora la trae reemplaza al registro incompleto (unica edicion permitida)
    completados = set(arribos_nuevos.loc[arribos_nuevos["fecha_llegada"].notna(), "hawb_key"])
    incompletos = arribos["hawb_key"].isin(completados) & arribos["fecha_llegada"].isna()
    if incompletos.any():
        log.info("   %d avisos sin fecha completados en esta corrida", int(incompletos.sum()))
        arribos = arribos[~incompletos]
    arribos, n_arr = anexar(arribos, arribos_nuevos, COLS_ARRIBOS)     # el reemplazo entra como fila nueva
    salidas, n_sal = anexar(salidas, salidas_nuevas, COLS_SALIDAS)
    notif, n_not = anexar(notif, notif_nuevas, COLS_NOTIF)

    kn = resolver(notif)
    vista = construir_seguimiento(bi, arribos, salidas, kn, corte_bi)
    controles_fwd(vista)
    embarques = construir_embarques(vista, hu, arribos)
    if dry_run:
        reporte_dry_run(vista, embarques, kn)
        log.info("DRY_RUN: no se escribio nada · quedaria +%d arribos · +%d salidas a DC · +%d correos FWD · "
                 "seguimiento: %s", n_arr, n_sal, n_not, vista["hito"].value_counts().to_dict())
        return

    if n_arr or n_sal or n_not or notif_nueva:
        if notif_nueva and C.RUTA_EVENTOS.exists():
            respaldar_historial()
        escribir_historial(arribos, salidas, notif)
        log.info("historial: +%d arribos (%d total) · +%d salidas a DC (%d total) · +%d correos FWD (%d total)",
                 n_arr, len(arribos), n_sal, len(salidas), n_not, len(notif))
    else:
        log.info("historial sin cambios")
    escribir_hojas({C.HOJA_VISTA: vista, C.HOJA_EMBARQUES: embarques}, mtime_bi)
    log.info("seguimiento: %s", vista["hito"].value_counts().to_dict())
    log.info("embarques: %d · bultos %s · %s kg · %s m3", len(embarques),
             f"{embarques['n_bultos'].sum():,.0f}", f"{embarques['peso_bruto_kg'].sum():,.1f}",
             f"{embarques['volumen_m3'].sum():,.2f}")
    log.info("=== FIN ===")


main = main_eventos
