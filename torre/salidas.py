# -*- coding: utf-8 -*-
"""
salidas.py — Consolidación sin cambios de lógica de las salidas y el cliente AIS:
  1. ais_buques.py          (posición AIS por nombre vía ProveedorA + ProveedorB Atlas, con freno)
  2. job_mail_arribados.py  (render HTML en paleta Cliente, Excel adjunto y envío/borrador Outlook) -> main_correo()
  3. job_radar_embarques.py (caché AIS catalogo_buques.json, mapa mundial y radar_embarques.html)  -> main_radar()

Redundancia R1 eliminada: main_correo() retorna el resultado de vista.armar() para pasarlo en memoria
a main_radar(armado=armado) dentro de la misma corrida.
"""

import html
import json
import logging
import math
import random
import re
import shutil
import tempfile
import time
import urllib.request
import webbrowser
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import numpy as np
import pandas as pd

import torre_config as C
import vista as V

# ══════════════════════════ CONFIG · AIS BUQUES ══════════════════════════════
TZ_PERU   = timezone(timedelta(hours=-5))
NAVEGADOR = "chrome"                 # perfil de curl_cffi, el mismo de tu script del globo
RE_COORD  = r'\\?"(?:{})\\?"\s*:\s*(-?\d+\.\d+)'

# ══════════════════════════ CONFIG · CORREO DE ARRIBADOS ═════════════════════
PARA           = []                  # ["alguien@cliente.com", ...] — obligatorio con DRY_RUN_CORREO = False
CC             = []
ASUNTO         = "Pedidos con mercadería arribada a Perú · {fecha}"
ADJUNTAR_EXCEL = True                # Excel con el detalle (embarques + lineas) para cruzar con otros sistemas
NOMBRE_EXCEL   = "mercaderia_arribada_{fecha}.xlsx"
MAX_PEDIDOS_ALERTA = 5               # pedidos que se nombran en cada alerta; el resto va como "y N más"

# Diseno para Outlook de escritorio (motor de Word): todo en tablas, estilos en linea y bgcolor; sin flex, grid,
# rgba ni imagenes. Paleta Cliente: rojo primario, morado conector, azul y neutros de soporte; nada de fondos
# grandes solo rojos y el texto chico en negro o blanco (alertas: texto negro sobre rojo palido con borde rojo)
ANCHO = 1200                         # como el correo M4
LETRA = "font-family:'Segoe UI',Calibri,Arial,sans-serif"
RED, RED_D2, PR2 = "#E4572E", "#A3341C", "#FCEDE8"               # Signature Red · Deep Red 2 · Pale Red 2
PUR2, PUR1, PP2  = "#5B3E96", "#3A2766", "#EEE9F7"               # Deep Purple 2 · Deep Purple 1 · Pale Purple 2
BLU2, PB2        = "#2F6DB5", "#EAEEF5"                          # Deep Blue 2 · Pale Blue 2
BLK, DG, MG, PG, W = "#1A1726", "#4A4658", "#A9A6B4", "#E6E2E4", "#FFFFFF"
CEBRA = "#F7F5F6"                    # Pale Gray aclarado: filas alternas
COLOR_ESTADO = {C.ARRIBADO: (PUR2, PP2), C.EN_RUTA: (BLU2, PB2)} # cabecera y banda de cada columna
TRAMO = {"hecho": PUR2, "en curso": BLU2, "en curso fuera de SLA": RED, "pendiente": PG}   # franja del recorrido
COLOR = {"banda": "#1f2a44"}         # encabezado del Excel adjunto (sin cambios)
MOT   = {"AIR": ("Aéreo", BLU2), "SEA": ("Marítimo", PUR1), "SEA-AIR": ("Marítimo-aéreo", DG)}

# ══════════════════════════ CONFIG · RADAR DE EMBARQUES ══════════════════════
LINK_DETALLE    = ""                 # «Copiar vínculo» de in_transit_peru.xlsx en OneDrive; vacío = la ruta local
URLS_MAPA       = ["https://cdn.jsdelivr.net/npm/world-atlas@2/countries-50m.json",
                   "https://unpkg.com/world-atlas@2/countries-50m.json"]
HORAS_CACHE_AIS = 6                  # posición más nueva que esto: no se vuelve a pedir
DIAS_REINTENTO  = 7                  # buque que ProveedorA no encontró: se vuelve a buscar a los 7 días
PAUSA_S         = (8, 15)            # espera al azar entre consultas
MAX_CONSULTAS   = 20                 # tope de consultas web por corrida (ProveedorA + ProveedorB)
CODIGOS_FRENO   = (403, 429, 503)
TIMEOUT_S       = 20
MARCA_DATOS, MARCA_MAPA = "/*__DATOS__*/{}", "/*__MAPA__*/null"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
log_correo = logging.getLogger("mail_arribados")
log_radar  = logging.getLogger("radar_embarques")


# ══════════════════════════ PARTE 1 · AIS BUQUES (ProveedorA + ProveedorB) ═══
def abrir_sesion():
    """Sesión de curl_cffi. Sin la librería: ImportError, y el radar sale sin posiciones (lo avisa)."""
    from curl_cffi import requests as cffi_requests
    return cffi_requests.Session(impersonate=NAVEGADOR)


class Web:
    """GET con freno. Después del primer bloqueo o del tope, todo get devuelve None y `frenado` dice por qué."""

    def __init__(self, sesion, pausa_s=(8, 15), maximo=20, codigos_freno=(403, 429, 503), timeout_s=20):
        self.sesion, self.pausa_s, self.maximo = sesion, pausa_s, maximo
        self.codigos_freno, self.timeout_s = set(codigos_freno), timeout_s
        self.consultas, self.frenado = 0, None

    def get(self, url):
        if self.frenado:
            return None
        if self.consultas >= self.maximo:
            self.frenado = f"tope de {self.maximo} consultas por corrida"
            return None
        if self.consultas:
            time.sleep(random.uniform(*self.pausa_s))
        self.consultas += 1
        try:
            resp = self.sesion.get(url, timeout=self.timeout_s)
        except Exception as e:                       # red caída o timeout: no es bloqueo, se sigue con el siguiente
            return _Fallo(f"{type(e).__name__} en {url.split('/')[2]}")
        if resp.status_code in self.codigos_freno:
            self.frenado = f"HTTP {resp.status_code} de {url.split('/')[2]}"
            return None
        return resp if resp.status_code == 200 else _Fallo(f"HTTP {resp.status_code} de {url.split('/')[2]}")


class _Fallo:
    """Respuesta que no sirve pero no frena la corrida (404, timeout)."""
    status_code, text = None, ""

    def __init__(self, motivo):
        self.motivo = motivo


def _ok(resp) -> bool:
    return resp is not None and not isinstance(resp, _Fallo)


def _motivo(web, resp, por_defecto):
    return web.frenado or (resp.motivo if isinstance(resp, _Fallo) else por_defecto)


def distancia_km(lat1, lon1, lat2, lon2) -> int:
    rad = math.pi / 180
    a = (math.sin((lat2 - lat1) * rad / 2) ** 2
         + math.cos(lat1 * rad) * math.cos(lat2 * rad) * math.sin((lon2 - lon1) * rad / 2) ** 2)
    return round(6371 * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))


def buscar_imo_mmsi(web, nombre: str) -> dict:
    """{'imo', 'mmsi', 'nombre_web', 'exacto'} o {'error'}. Entre los resultados manda el que se llama exactamente
    igual; si ninguno, el primero con exacto = False (posible homónimo: el radar lo marca)."""
    from bs4 import BeautifulSoup
    resp = web.get(f"https://ais-proveedor-a.example/vessels?name={quote(nombre)}")
    if not _ok(resp):
        return {"error": _motivo(web, resp, "ProveedorA no respondió")}
    enlaces = BeautifulSoup(resp.text, "html.parser").select("a.ship-link")
    if not enlaces:
        return {"error": "ProveedorA no encontró el buque"}
    plano = lambda t: re.sub(r"\s+", " ", t or "").strip().upper()
    nombre_de = lambda a: plano((a.select_one(".slna") or a).get_text(" "))     # .slna: el nombre en la lista
    exactos = [a for a in enlaces if nombre_de(a) == nombre]                     # «MSC ORSOLA» antes que «MSC ORSOLA II»
    enlace = (exactos or enlaces)[0]
    imo = re.search(r"/details/(\d+)", enlace.get("href", ""))
    detalle = web.get(f"https://ais-proveedor-a.example{enlace.get('href', '')}")
    mmsi = None
    if _ok(detalle):
        sopa = BeautifulSoup(detalle.text, "html.parser")
        djson = sopa.find(id="djson")
        if djson is not None and djson.has_attr("data-json"):
            try:
                mmsi = str(json.loads(djson["data-json"]).get("mmsi") or "") or None
            except ValueError:
                pass
        if not mmsi:
            hallado = re.search(r"\b\d{7}\s*/\s*(\d{9})\b", sopa.get_text(" "))     # «IMO / MMSI» del detalle
            mmsi = hallado.group(1) if hallado else None
    if not (imo and mmsi):
        return {"error": _motivo(web, detalle, "ProveedorA sin IMO/MMSI para el buque")}
    return {"imo": imo.group(1), "mmsi": mmsi, "nombre_web": nombre_de(enlace), "exacto": bool(exactos)}


def _a_hora_peru(iso):
    if not isinstance(iso, str) or not iso:
        return None
    try:
        return datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone(TZ_PERU).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return None


def _extraer_bloque_json(html_txt: str, clave: str):
    """Objeto {...} balanceado que sigue a "clave": en el payload Next.js (escapado o no)."""
    hallado = re.search(rf'(\\)?"{clave}(\\)?"\s*:\s*{{', html_txt)
    if not hallado:
        return None
    escapado, inicio, nivel = bool(hallado.group(1)), hallado.end() - 1, 0
    for i in range(inicio, len(html_txt)):
        if html_txt[i] == "{":
            nivel += 1
        elif html_txt[i] == "}":
            nivel -= 1
            if nivel == 0:
                crudo = html_txt[inicio:i + 1]
                if escapado:
                    crudo = crudo.replace('\\"', '"').replace("\\\\", "\\")
                try:
                    return json.loads(crudo)
                except ValueError:
                    return None
    return None


def _campo(html_txt: str, clave: str):
    hallado = re.search(rf'\\?"{clave}\\?"\s*:\s*\\?"([^"\\]+)\\?"', html_txt)
    return hallado.group(1).strip() if hallado else None


def _numero(html_txt: str, clave: str):
    hallado = re.search(rf'\\?"{clave}\\?"\s*:\s*(-?\d+(?:\.\d+)?)', html_txt)
    return float(hallado.group(1)) if hallado else None


def posicion_proveedor_b(web, nombre: str, imo: str, mmsi: str, puerto: str) -> dict:
    """Posición, ETA al `puerto` (UN/LOCODE) y escalas, o {'error', 'url'}. Igual que tu script del globo:
    la primera latitud/longitud del payload es la del buque; la ETA AIS vale si su destino es el puerto, si no
    manda la ETA del itinerario para ese puerto. Arribado = la última escala es el puerto y no tiene zarpe."""
    slug = re.sub(r"[^a-z0-9]+", "-", nombre.lower()).strip("-")
    url = f"https://ais-proveedor-b.example/vessel/imo:{imo}/mmsi:{mmsi}/name:{slug}"
    resp = web.get(url)
    if not _ok(resp):
        return {"error": _motivo(web, resp, "ProveedorB no respondió"), "url": url}
    html_txt = resp.text
    lats = re.findall(RE_COORD.format("latitude|lat"), html_txt)
    lons = re.findall(RE_COORD.format("longitude|lng|lon"), html_txt)
    if not (lats and lons):
        return {"error": "ProveedorB sin coordenadas en la página", "url": url}
    lat, lon = float(lats[0]), float(lons[0])
    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return {"error": f"coordenadas fuera de rango ({lat}, {lon})", "url": url}

    pasados = (_extraer_bloque_json(html_txt, "vesselPortCalls") or {}).get("data") or []
    escalas_pasadas = [{"puerto": p.get("portName") or p.get("portUnlocode"), "locode": p.get("portUnlocode"),
                        "arribo": _a_hora_peru(p.get("arrivalTime")), "zarpe": _a_hora_peru(p.get("departureTime"))}
                       for p in pasados if isinstance(p, dict)]
    futuros = _extraer_bloque_json(html_txt, "futurePortCalls") or {}
    escalas_futuras, eta_itinerario = [], None
    for f in futuros.get("future_port_calls") or []:
        if not isinstance(f, dict):
            continue
        nombre_puerto = (f.get("port_name") or f.get("port_unlocode") or f.get("port_code") or f.get("unlocode")
                         or f.get("port") or (futuros.get("next_port") if f.get("is_next_port") else None))
        eta = f"{f.get('eta') or ''} {(f.get('eta_time') or '')[:5]}".strip() or None
        escalas_futuras.append({"puerto": nombre_puerto, "eta": eta, "es_proximo": bool(f.get("is_next_port"))})
        es_puerto = puerto in json.dumps(f).upper() or (f.get("is_next_port") and futuros.get("next_port") == puerto)
        if eta_itinerario is None and eta and es_puerto:
            eta_itinerario = eta

    destino, eta_ais = _campo(html_txt, "destinationUnlocode"), _a_hora_peru(_campo(html_txt, "eta"))
    teu = _numero(html_txt, "teu")
    return {"lat": lat, "lon": lon, "url": url,
            "actualizado": _a_hora_peru(_campo(html_txt, "lastUpdatedAt") or _campo(html_txt, "updatedAt")),
            "bandera": _campo(html_txt, "flag"), "teu": int(teu) if teu else None, "calado": _numero(html_txt, "draught"),
            "puerto_anterior": futuros.get("current_port") or (escalas_pasadas[0]["locode"] if escalas_pasadas else None),
            "proximo_puerto": futuros.get("next_port") or destino,
            "eta_puerto": eta_ais if destino == puerto and eta_ais else eta_itinerario,
            "arribado": bool(escalas_pasadas and escalas_pasadas[0]["locode"] == puerto
                             and not escalas_pasadas[0]["zarpe"]),
            "escalas_pasadas": escalas_pasadas[:4], "escalas_futuras": escalas_futuras[:4]}


# ══════════════════════════ PARTE 2 · EXCEL Y CORREO ARRIBADOS ═══════════════
def _sold_to_de(pedidos, sold: dict):
    """'4330001 | 4330002' (como lo une por_linea) -> sus sold-to, en el mismo orden."""
    if pd.isna(pedidos):
        return pd.NA
    return " | ".join(sold.get(p.strip(), "—") for p in str(pedidos).split("|"))


def hojas_excel(tabla: pd.DataFrame, lineas: pd.DataFrame) -> dict:
    """Dos hojas con nombres de columna legibles: resumen por embarque y detalle por linea del bi."""
    hoy = pd.Timestamp.now().normalize()
    por_emb = tabla.drop_duplicates("embarque").set_index("embarque")
    sold = {p: s for lista, sts in zip(tabla["pedidos"], tabla["sold_to"]) for (p, _), s in zip(lista, sts) if s}
    b2b = tabla.groupby("embarque")["cantidad"].sum()
    pares = tabla.groupby("embarque")["pedidos"].agg(lambda s: {par for lista in s for par in lista})
    embarques = pd.DataFrame({
        "Embarque": por_emb.index,
        "Estado": por_emb["estado"].values,
        "Transporte": [MOT.get(str(m), (str(m).title(),))[0] for m in por_emb["mot"]],
        "Llegada": por_emb["llegada"].dt.normalize().values,
        "Salida a DC": por_emb["salida"].dt.normalize().where(por_emb["estado"].eq(C.EN_RUTA)).values,
        "Hora salida": por_emb["salida_hora"].where(por_emb["estado"].eq(C.EN_RUTA)).values,
        "Días desde llegada": (hoy - por_emb["llegada"].dt.normalize()).dt.days.values,
        "Unidades total": por_emb["unidades_embarque"].values,
        "Unidades B2B": b2b.reindex(por_emb.index).values,
        "Unidades stock/otros": (por_emb["unidades_embarque"] - b2b.reindex(por_emb.index)).values,
        "Bultos (Lake)": por_emb["bultos_embarque"].values,
        "Kg (Lake)": por_emb["kg_embarque"].values,
        "m³ (Lake)": por_emb["m3_embarque"].values,
        "Bultos aviso": por_emb.get("aviso_total", pd.Series(index=por_emb.index, dtype=float)).values,
        "Buenos aviso": por_emb.get("aviso_buenos", pd.Series(index=por_emb.index, dtype=float)).values,
        "Malos aviso": por_emb.get("aviso_malos", pd.Series(index=por_emb.index, dtype=float)).values,
        "Faltantes aviso": por_emb.get("aviso_faltantes", pd.Series(index=por_emb.index, dtype=float)).values,
        "Pedidos B2B": [len({p for p, _ in pares[e]}) for e in por_emb.index],
        "Clientes": [len({c for _, c in pares[e]}) for e in por_emb.index],
    })
    estado = lineas["bl_awb_num"].map(por_emb["estado"])                 # el estado es del HAWB, no de la linea
    detalle = pd.DataFrame({
        "Embarque": lineas["bl_awb_num"],
        "Estado": estado,
        "Transporte": [MOT.get(str(m).upper(), (str(m).title(),))[0] for m in lineas["mot"]],
        "Llegada": pd.to_datetime(lineas["llegada_fecha"], errors="coerce").dt.normalize(),
        "Salida a DC": lineas["bl_awb_num"].map(por_emb["salida"]).dt.normalize().where(estado.eq(C.EN_RUTA)),
        "SO tránsito": lineas["so"],
        "MTM": lineas["mtm"],
        "Cantidad": pd.to_numeric(lineas["transit_qty"], errors="coerce"),
        "DN": lineas.get("delivery_order_number"),
        "Vínculo": lineas["vinculo"],
        "Pedido (433)": lineas["so_local"],
        "Cliente": lineas["cliente"],
        "Sold-to": lineas["so_local"].map(lambda v: _sold_to_de(v, sold)),   # 'a | b': pocos cientos de filas
        "N° pedidos": lineas["n_pedidos"],
        "Segmento": lineas.get("segment"),
        "Key": lineas["key"],
    })
    orden_vinculo = {"con pedido": 0, "sin pedido (stock)": 1, "SO en OrderMatch con otro MTM": 2}   # B2B primero
    detalle = (detalle.assign(_e=detalle["Estado"].map({C.ARRIBADO: 0, C.EN_RUTA: 1}),
                              _v=detalle["Vínculo"].map(orden_vinculo))
                      .sort_values(["_e", "Llegada", "Embarque", "_v", "SO tránsito", "MTM"], na_position="last")
                      .drop(columns=["_e", "_v"]))
    return {"embarques": embarques, "lineas": detalle}


def escribir_excel(hojas: dict, ruta: Path) -> None:
    """Formato para leer y filtrar: encabezado oscuro, filtro, primera fila fija, anchos y formatos de numero."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    formatos = {"Llegada": "dd/mm/yyyy", "Salida a DC": "dd/mm/yyyy", "Kg (Lake)": "#,##0", "m³ (Lake)": "#,##0.00"}
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
            ws = w.sheets[nombre]
            for celda in ws[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill("solid", fgColor=COLOR["banda"].lstrip("#"))
                celda.alignment = Alignment(vertical="center", wrap_text=True)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, start=1):
                largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(500) if pd.notna(v)])
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, largo + 2), 45)
                formato = formatos.get(col, "#,##0" if pd.api.types.is_numeric_dtype(df[col]) else None)
                if formato:
                    for fila in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                        fila[0].number_format = formato


# Piezas: cabecera negra con cifras -> «Requiere accion» -> columnas Arribado | En ruta al DC -> definiciones.
def _esc(v) -> str:
    return html.escape(str(v))


def _num(x, dec=0) -> str:
    return "—" if pd.isna(x) else f"{x:,.{dec}f}"


def _plural(n, singular: str, plural: str = None) -> str:
    return f"{_num(n)} {singular if n == 1 else (plural or singular + 's')}"


def _tabla(filas: str, ancho="100%", attrs="") -> str:
    ancho = f' width="{ancho}"' if ancho else ""
    return f'<table role="presentation"{ancho} cellpadding="0" cellspacing="0" border="0"{attrs}>{filas}</table>'


def _franja(color: str, alto: int) -> str:
    """Fila de color con alto fijo: va en atributo y en line-height para que Outlook no la engorde."""
    return (f'<tr><td height="{alto}" bgcolor="{color}" style="height:{alto}px;font-size:{alto}px;'
            f'line-height:{alto}px;mso-line-height-rule:exactly">&nbsp;</td></tr>')


def _pastilla(texto: str, fondo: str) -> str:
    return (f'<span style="{LETRA};background:{fondo};color:{W};font-size:10px;font-weight:700;padding:2px 4px;'
            f'letter-spacing:.3px;white-space:nowrap">&nbsp;{texto}&nbsp;</span>')


def _pastilla_mot(mot) -> str:
    texto, color = MOT.get(str(mot), (str(mot).title() or "—", DG))
    return _pastilla(_esc(texto.upper()), color)


def _cifra(valor, etiqueta: str, color=BLK, color_etiqueta=DG, tamano=34, pad="0 22px 0 0") -> str:
    return (f'<td valign="bottom" style="padding:{pad}"><div style="{LETRA};font-size:{tamano}px;font-weight:700;'
            f'line-height:{tamano + 2}px;color:{color}">{valor}</div><div style="{LETRA};font-size:10px;'
            f'color:{color_etiqueta};text-transform:uppercase;letter-spacing:.6px;white-space:nowrap">{etiqueta}</div></td>')


def _th(texto: str, ancho: str, color: str, alinear="left") -> str:
    return (f'<th align="{alinear}" width="{ancho}" style="{LETRA};font-size:10px;font-weight:700;color:{DG};'
            f'text-transform:uppercase;letter-spacing:.4px;padding:6px 8px;border-bottom:2px solid {color};'
            f'white-space:nowrap">{texto}</th>')


def _td(texto: str, fila: int, alinear="left", fuerte=False) -> str:
    cebra = f' bgcolor="{CEBRA}"' if fila % 2 else ""
    return (f'<td align="{alinear}" valign="top"{cebra} style="{LETRA};font-size:12px;color:{BLK};'
            f'{"font-weight:700;" if fuerte else ""}padding:6px 8px;border-bottom:1px solid {PG}">{texto}</td>')


def _relativo(fecha, hoy) -> str:
    if pd.isna(fecha):
        return ""
    dias_dif = (hoy - fecha.normalize()).days
    if dias_dif in (-1, 0, 1):
        return {-1: "mañana", 0: "hoy", 1: "ayer"}[dias_dif]
    return f"hace {_plural(dias_dif, 'día')}" if dias_dif > 0 else f"en {_plural(-dias_dif, 'día')}"


def _salida_txt(g0, formato: str) -> str:
    hora = f' {_esc(g0["salida_hora"])}' if pd.notna(g0["salida_hora"]) else ""
    return "sin fecha" if pd.isna(g0["salida"]) else f'{g0["salida"]:{formato}}{hora}'


def _cuando(g0, hoy) -> str:
    """La fecha que importa segun el estado: la salida a DC si FWD ya la programo; si no, la llegada a Peru (sin
    'hace N dias': al lado del conteo en dias habiles del SLA se leeria como otra cifra)."""
    if g0["estado"] == C.EN_RUTA:
        verbo = "Salió" if pd.notna(g0["salida"]) and g0["salida"].normalize() < hoy else "Sale"
        relativo = _relativo(g0["salida"], hoy)
        return (f'{verbo} a DC <b style="color:{BLK}">{_salida_txt(g0, "%d/%m/%Y")}</b>'
                + (f" · {relativo}" if relativo else ""))
    llegada = g0["llegada_embarque"]
    eta = " (por ETA)" if g0["llegada_fuente"] == "ETA vencida" else ""
    return f'Llegó <b style="color:{BLK}">{"sin fecha" if pd.isna(llegada) else f"{llegada:%d/%m/%Y}"}</b>{eta}'


def _alertas(g0) -> tuple:
    """Motivos de «Requiere accion» del embarque y el siguiente paso de cada uno."""
    motivos, pasos = [], []
    if g0["fuera_sla"]:
        transporte = MOT.get(str(g0["mot"]), (str(g0["mot"]),))[0].lower()
        motivos.append(f'Fuera de SLA · {g0["dias_hab"]:.0f} días hábiles sin salida a DC ({transporte}: {g0["sla_hab"]:.0f})')
        pasos.append(C.PASO_FUERA_SLA)
    total, malos, faltan = (pd.to_numeric(g0[c], errors="coerce") for c in ("aviso_total", "aviso_malos", "aviso_faltantes"))
    partes = []
    if pd.notna(malos) and malos > 0:
        partes.append(f"{_num(malos)} de {_num(total)} bultos malos" if pd.notna(total) else _plural(malos, "bulto malo", "bultos malos"))
    if pd.notna(faltan) and faltan > 0:
        partes.append(_plural(faltan, "faltante"))
    if partes:
        motivos.append("Aviso de llegada: " + " · ".join(partes))
        pasos.append(C.PASO_AVISO)
    return motivos, pasos


def _recorrido(g0, hoy) -> str:
    """Franja fina de 4 tramos: salida de origen -> arribo a Peru -> salida a DC -> ingreso al DC."""
    dia = lambda f: "—" if pd.isna(f) else f"<b>{f:%d/%m}</b>"
    arribo = dia(g0["llegada_embarque"]) + (" · por ETA" if g0["llegada_fuente"] == "ETA vencida" else "")
    if g0["estado"] == C.EN_RUTA:
        relativo = _relativo(g0["salida"], hoy)
        salida = f'<b>{_salida_txt(g0, "%d/%m")}</b>' + (f" · {relativo}" if relativo else "")
        salio = pd.notna(g0["salida"]) and g0["salida"].normalize() < hoy
        estados = ["hecho", "hecho", "hecho", "en curso"] if salio else ["hecho", "hecho", "en curso", "pendiente"]
        ingreso = "en tránsito" if salio else "—"
    else:
        salida = "por programar" + ("" if pd.isna(g0["dias_hab"]) else f' · <b>{g0["dias_hab"]:.0f} d háb.</b>')
        estados = ["hecho", "hecho", "en curso fuera de SLA" if g0["fuera_sla"] else "en curso", "pendiente"]
        ingreso = "—"
    tramos = [("Salida de origen", dia(g0["atd_origen"]), 22), ("Arribo a Perú", arribo, 24),
              ("Salida a DC", salida, 32), ("Ingreso al DC", ingreso, 22)]
    celdas = ""
    for i, ((nombre, valor, ancho), estado) in enumerate(zip(tramos, estados)):
        actual = estado.startswith("en curso")
        celdas += (f'<td width="{ancho}%" valign="top" style="padding:0 {4 if i < 3 else 0}px 0 0">' + _tabla(
            _franja(TRAMO[estado], 4)
            + f'<tr><td style="{LETRA};font-size:10px;color:{BLK if actual else DG};{"font-weight:700;" if actual else ""}'
              f'padding-top:4px;white-space:nowrap">{nombre}</td></tr>'
            + f'<tr><td style="{LETRA};font-size:11px;color:{BLK}">{valor}</td></tr>') + '</td>')
    return _tabla(f"<tr>{celdas}</tr>")


def _bloque(embarque, grupo: pd.DataFrame, color: str, hoy, accion: bool) -> str:
    """Un embarque: id, transporte y volumen completo; su recorrido; y sus lineas con pedido B2B."""
    g0 = grupo.iloc[0]
    b2b = pd.to_numeric(grupo["cantidad"], errors="coerce").sum()
    stock = g0["unidades_embarque"] - b2b if pd.notna(g0["unidades_embarque"]) else float("nan")
    unidades = (f'{_num(g0["unidades_embarque"])} u. sin pedido' if not b2b
                else f'{_num(g0["unidades_embarque"])} u. (B2B {_num(b2b)} · stock {_num(stock)})'
                if pd.notna(stock) and stock > 0 else f"{_num(b2b)} u. B2B")
    bultos = f'{_num(g0["bultos_embarque"])} bulto{"" if g0["bultos_embarque"] == 1 else "s"}'
    sin_lake = (f"{bultos} según aviso · " if pd.notna(g0["bultos_embarque"]) else "") + unidades
    volumen = (f"{sin_lake} · fuera del in-transit" if g0["fuera_in_transit"]                         # sin Lake
               else f'{bultos} · {_num(g0["kg_embarque"])} kg · {_num(g0["m3_embarque"], 2)} m&#179; · {unidades}')
    marca = f'&nbsp;{_pastilla("ACCIÓN", RED)}' if accion else ""          # el detalle esta en «Requiere accion»
    cabeza = (f'<tr><td valign="bottom" style="{LETRA};font-size:15px;font-weight:700;color:{BLK};padding:0 8px 8px 0;'
              f'white-space:nowrap">{_esc(embarque)}&nbsp;&nbsp;{_pastilla_mot(g0["mot"])}{marca}</td>'
              f'<td align="right" valign="bottom" style="{LETRA};font-size:11px;color:{DG};padding:0 0 8px">{volumen}</td></tr>')
    cabecera = (_th("SO tránsito", "17%", color) + _th("MTM", "19%", color) + _th("Cant.", "10%", color, "right")
                + _th("Pedido", "19%", color) + _th("Cliente", "35%", color))
    filas = "".join("<tr>" + _td(_esc(r.so), i) + _td(_esc(r.mtm), i)
                    + _td(_num(r.cantidad if r.pedidos else getattr(r, "unid_stock", np.nan)), i, "right", True)
                    + _td("<br>".join(_esc(p) for p, _ in r.pedidos) or "—", i)     # 433 y cliente alineados
                    + _td("<br>".join(_esc(c) + (f' <span style="color:{DG};white-space:nowrap">·&nbsp;Sold&nbsp;to&nbsp;{_esc(s)}</span>' if s else "")
                                      for (_, c), s in zip(r.pedidos, r.sold_to))   # sold-to sin cortes: baja entero
                          or "Sin pedido en OrderMatch (stock)", i) + "</tr>"
                    for i, r in enumerate(grupo.itertuples()))
    return ('<tr><td style="padding:18px 0 0">' + _tabla(
        f'<tr><td>{_tabla(cabeza)}</td></tr><tr><td>{_recorrido(g0, hoy)}</td></tr>'
        f'<tr><td style="padding-top:10px">{_tabla(f"<tr>{cabecera}</tr>{filas}")}</td></tr>') + '</td></tr>')


def _columna(sub: pd.DataFrame, estado: str, titulo: str, nota: str, hoy, marcados: set) -> str:
    """Una columna (Arribado o En ruta al DC): cabecera de color, banda con cifras y un bloque por embarque."""
    color, palido = COLOR_ESTADO[estado]
    n_ped = len({p for lista in sub["pedidos"] for p, _ in lista})
    fuera = int(sub.drop_duplicates("embarque")["fuera_sla"].sum())
    franja = _tabla("<tr>" + _cifra(sub["embarque"].nunique(), "Embarques", tamano=30) + _cifra(n_ped, "Pedidos", tamano=30)
                    + _cifra(_num(pd.to_numeric(sub["cantidad"], errors="coerce").sum()), "Unid. B2B", tamano=30)
                    + (_cifra(fuera, "Fuera de SLA", RED if fuera else BLK, tamano=30) if estado == C.ARRIBADO else "")
                    + "</tr>", ancho=None)
    bloques = "".join(_bloque(emb, grupo, color, hoy, emb in marcados) for emb, grupo in sub.groupby("embarque", sort=False))
    bloques = bloques or (f'<tr><td style="{LETRA};font-size:12px;color:{BLK};padding:16px 4px">'
                          f'Ningún embarque en este estado.</td></tr>')
    return ('<td width="50%" valign="top" style="padding:0 10px">' + _tabla(
        f'<tr><td bgcolor="{color}" style="padding:12px 14px"><div style="{LETRA};font-size:17px;font-weight:700;'
        f'color:{W}">{titulo}</div><div style="{LETRA};font-size:11px;color:{W}">{nota}</div></td></tr>'
        f'<tr><td bgcolor="{palido}" style="padding:12px 14px">{franja}</td></tr>{bloques}') + '</td>')


def _requiere_accion(tabla: pd.DataFrame, hoy) -> tuple:
    """Bloque de arriba: un renglon por embarque con alerta (motivo, siguiente paso y pedidos que afecta), el mas
    atrasado contra su SLA primero. Devuelve el HTML y los embarques que entraron (se marcan en su columna)."""
    emb = tabla.drop_duplicates("embarque").set_index("embarque")
    emb = (emb.assign(_exceso=(emb["dias_hab"] - emb["sla_hab"]).where(emb["fuera_sla"]))
              .sort_values(["_exceso", "llegada_embarque"], ascending=[False, True], na_position="last"))
    filas, marcados = [], []
    for embarque, g0 in emb.iterrows():                      # un renglon por embarque del dia: pocos
        motivos, pasos = _alertas(g0)
        if not motivos:
            continue
        marcados.append(embarque)
        grupo = tabla[tabla["embarque"].eq(embarque)]
        pares = list(dict.fromkeys(par for lista in grupo["pedidos"] for par in lista))
        nombrados = " · ".join(f"<b>{_esc(p)}</b> {_esc(c)}" for p, c in pares[:MAX_PEDIDOS_ALERTA])
        resto = len(pares) - MAX_PEDIDOS_ALERTA
        afecta = (f'{_plural(len(pares), "pedido")} · {_num(pd.to_numeric(grupo["cantidad"], errors="coerce").sum())} '
                  f'u. B2B: {nombrados}' + (f" · y {_plural(resto, 'pedido')} más" if resto > 0 else "")
                  if pares else f'Sin pedido en OrderMatch · {_num(g0["unidades_embarque"])} u. de stock')
        filas.append(
            '<tr><td style="padding:0 0 10px">' + _tabla(
                f'<tr><td bgcolor="{PR2}" style="padding:12px 16px;border-left:6px solid {RED}">' + _tabla(
                    f'<tr><td valign="top" style="{LETRA};font-size:15px;font-weight:700;color:{RED_D2}">{" · ".join(motivos)}</td>'
                    f'<td rowspan="3" align="right" valign="top" style="padding-left:20px;white-space:nowrap">'
                    f'<div style="{LETRA};font-size:15px;font-weight:700;color:{BLK}">{_esc(embarque)}&nbsp; {_pastilla_mot(g0["mot"])}</div>'
                    f'<div style="{LETRA};font-size:12px;color:{DG}">{_cuando(g0, hoy)}</div></td></tr>'
                    f'<tr><td style="{LETRA};font-size:12px;color:{BLK};padding-top:3px"><b>Siguiente paso:</b> {" y ".join(pasos)}</td></tr>'
                    f'<tr><td style="{LETRA};font-size:11px;color:{BLK};padding-top:3px">{afecta}</td></tr>')
                + '</td></tr>') + '</td></tr>')
    if not filas:
        return "", set()
    titulo = ('<tr><td style="padding:0 0 10px">' + _tabla(
        f'<tr><td width="5" bgcolor="{RED}" style="width:5px;font-size:1px;line-height:1px">&nbsp;</td>'
        f'<td style="padding:2px 12px"><span style="{LETRA};font-size:17px;font-weight:700;color:{BLK}">Requiere acción</span>'
        f'<span style="{LETRA};font-size:12px;color:{DG}">&nbsp;&nbsp;{_plural(len(filas), "embarque")} · '
        f'con su siguiente paso</span></td></tr>') + '</td></tr>')
    return titulo + "".join(filas), set(marcados)


def cuerpo_html(tabla: pd.DataFrame, corte) -> str:
    hoy, desde = pd.Timestamp.now().normalize(), V.inicio_en_ruta()
    pares = {par for lista in tabla["pedidos"] for par in lista}
    n_emb, n_ped, n_cli = tabla["embarque"].nunique(), len({p for p, _ in pares}), len({x for _, x in pares})
    por_emb = tabla.drop_duplicates("embarque")
    total = lambda col: pd.to_numeric(por_emb[col], errors="coerce").sum(min_count=1)
    u_b2b, u_total = pd.to_numeric(tabla["cantidad"], errors="coerce").sum(), total("unidades_embarque")
    accion, marcados = _requiere_accion(tabla, hoy)

    kpis = (_cifra(n_emb, "Embarques", W, MG) + _cifra(n_ped, "Pedidos", W, MG) + _cifra(n_cli, "Clientes", W, MG)
            + _cifra(_num(u_b2b), f"Unid. B2B · de {_num(u_total)}", W, MG)
            + _cifra(len(marcados), "Requieren acción", RED if marcados else W, MG, pad="0"))
    cabecera = (f'<tr><td bgcolor="{BLK}" style="padding:24px 28px 20px">' + _tabla(
        f'<tr><td valign="bottom"><div style="{LETRA};font-size:11px;font-weight:700;color:{MG};letter-spacing:1.2px">'
        f'IN-TRANSIT PERÚ · MERCADERÍA ARRIBADA</div><div style="{LETRA};font-size:25px;font-weight:700;color:{W};'
        f'padding-top:2px">Pedidos de clientes con carga en Perú</div><div style="{LETRA};font-size:12px;color:{MG};'
        f'padding-top:4px">Corte {corte:%d/%m/%Y %H:%M} · {_num(total("bultos_embarque"))} bultos · '
        f'{_num(total("m3_embarque"), 1)} m&#179; · {_num(total("kg_embarque"))} kg</div></td>'
        f'<td align="right" valign="bottom">{_tabla("<tr>" + kpis + "</tr>", ancho=None)}</td></tr>')
        + '</td></tr>' + _franja(PUR2, 4))

    columnas = (_columna(tabla[tabla["estado"].eq(C.ARRIBADO)], C.ARRIBADO, "Arribado",
                         "Pendiente de salida a DC · más antiguos primero", hoy, marcados)
                + _columna(tabla[tabla["estado"].eq(C.EN_RUTA)], C.EN_RUTA, "En ruta al DC",
                           f"FWD programó la salida (desde el {desde:%d/%m}) · el DC la recibe ese día", hoy, marcados))
    cuerpo = ('<tr><td style="padding:22px 18px 18px">' + _tabla(
        (f'<tr><td colspan="2" style="padding:0 10px 12px">{_tabla(accion)}</td></tr>' if accion else "")
        + f"<tr>{columnas}</tr>") + '</td></tr>')

    leyenda = " &nbsp;".join(f'<span style="background:{c};font-size:9px">&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;{t}'
                             for t, c in TRAMO.items())
    sla = " · ".join(f"{MOT[m][0].lower()} {d}" for m, d in C.SLA_ARRIBO_HAB.items() if m in MOT)
    pie = (f'<tr><td style="{LETRA};font-size:10px;line-height:15px;color:{DG};padding:16px 28px;border-top:1px solid {PG}">'
           f"<b>Requiere acción:</b> arribado fuera de SLA, o bultos malos o faltantes según el aviso del agente. "
           f"<b>SLA de arribo:</b> días hábiles (lun–vie) desde la llegada sin salida a DC: {sla}. "
           f"<b>Llegada:</b> aéreo, fecha del aviso de llegada (PDF del agente; si no es legible, la del correo); marítimo, "
           f"ATA de los correos de FWD o del reporte y, si no llega, su ETA vigente ya vencida («por ETA»). "
           f"<b>Recorrido:</b> {leyenda}; salida de origen = ATD de los correos de FWD o, si no, del reporte. "
           f"<b>Cant.:</b> unidades en tránsito de la línea con pedido B2B; un embarque sin ningún pedido va con "
           f"una fila de su stock. <b>Volumen:</b> bultos, kg y m&#179; de Lake "
           f"y unidades de todo el HAWB/BL, incluido stock y e-commerce sin pedido. <b>Fuera del in-transit:</b> lo que el "
           f"reporte ya quitó (aéreos al llegar al terminal, marítimos antes de su salida a DC); sus líneas salen del "
           f"historial de tránsito, sus bultos del aviso y su pedido de la última copia del OrderMatch. "
           f"<b>Pedido y cliente:</b> OrderMatch, por SO de tránsito y MTM; el sold-to es el del pedido (433). <b>Salida a DC:</b> programación de FWD; con ella "
           f"el embarque pasa a En ruta al DC y sale de la lista "
           f"{_plural(C.DIAS_HABILES_ENTREGA, 'día hábil', 'días hábiles')} después de su salida. "
           f"{'Detalle por línea en el Excel adjunto. ' if ADJUNTAR_EXCEL else ''}Correo generado automáticamente.</td></tr>")

    previa = ((f"{len(marcados)} requiere{'n' if len(marcados) != 1 else ''} acción · " if marcados else "")
              + f"{_plural(n_emb, 'embarque')} · {_plural(n_ped, 'pedido')} · {_plural(n_cli, 'cliente')} con carga en Perú")
    return (f'<div style="display:none;max-height:0;overflow:hidden;mso-hide:all">{previa}</div>'
            f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="{PG}">'
            f'<tr><td align="center" style="padding:20px 0">'
            + _tabla(cabecera + cuerpo + pie, ancho=ANCHO, attrs=f' bgcolor="{W}" style="width:{ANCHO}px"')
            + '</td></tr></table>')


def enviar(asunto: str, cuerpo: str, adjunto: Path = None) -> None:
    assert C.DRY_RUN_CORREO or PARA, "DRY_RUN_CORREO = False sin destinatarios en PARA"
    import win32com.client as win32       # aca y no arriba: el resto del modulo no depende de Outlook

    mail = win32.Dispatch("Outlook.Application").CreateItem(0)
    mail.To, mail.CC, mail.Subject, mail.HTMLBody = "; ".join(PARA), "; ".join(CC), asunto, cuerpo
    if adjunto is not None:
        mail.Attachments.Add(str(adjunto))          # Outlook copia el archivo al correo: el temporal se puede borrar
    if C.DRY_RUN_CORREO:
        mail.Display(False)
        log_correo.info("DRY_RUN: borrador abierto en Outlook — no se envio")
    else:
        mail.Send()
        log_correo.info("enviado a %s", "; ".join(PARA))


def main_correo(armado=None):
    """Arma, formatea y envía (o abre borrador) el correo de arribados. Devuelve `armado` = (tabla, corte, lineas)
    para que el radar reutilice la tabla en memoria (R1) sin volver a ejecutar armar()."""
    log_correo.info("=== CORREO · aviso de llegada -> HAWB -> SO de transito -> 433 + cliente ===")
    tabla, corte, lineas = V.armar() if armado is None else armado
    if tabla.empty:
        log_correo.info("ningun pedido de cliente arribado ni en ruta al DC — no se arma correo")
        return tabla, corte, lineas
    pares = {par for lista in tabla["pedidos"] for par in lista}
    log_correo.info("correo: %d embarques · %d lineas · %d pedidos · %d clientes · %s",
                    tabla["embarque"].nunique(), len(tabla),
                    len({p for p, _ in pares}), len({c for _, c in pares}),
                    tabla.drop_duplicates("embarque")["estado"].value_counts().to_dict())
    ahora = pd.Timestamp.now()
    asunto, cuerpo = ASUNTO.format(fecha=f"{ahora:%d/%m/%Y}"), cuerpo_html(tabla, corte)
    if not ADJUNTAR_EXCEL:
        enviar(asunto, cuerpo)
    else:
        carpeta = Path(tempfile.mkdtemp(prefix="mail_arribados_"))
        try:
            ruta = carpeta / NOMBRE_EXCEL.format(fecha=f"{ahora:%Y%m%d_%H%M}")
            hojas = hojas_excel(tabla, lineas)
            escribir_excel(hojas, ruta)
            log_correo.info("adjunto %s: %d embarques · %d lineas",
                            ruta.name, len(hojas["embarques"]), len(hojas["lineas"]))
            enviar(asunto, cuerpo, ruta)
        finally:
            shutil.rmtree(carpeta, ignore_errors=True)   # no queda otro Excel en disco
    log_correo.info("=== FIN ===")
    return tabla, corte, lineas


# ══════════════════════════ PARTE 3 · RADAR HTML Y CACHÉ AIS ═════════════════
def leer_catalogo() -> dict:
    if not C.RUTA_CATALOGO.exists():
        return {}
    try:
        return json.loads(C.RUTA_CATALOGO.read_text(encoding="utf-8"))
    except ValueError as e:                    # es caché: se rehace, pero se avisa
        log_radar.warning("%s ilegible (%s): se empieza de cero", C.RUTA_CATALOGO.name, e)
        return {}


def guardar_catalogo(catalogo: dict) -> None:
    tmp = C.RUTA_CATALOGO.with_suffix(".tmp")
    tmp.write_text(json.dumps(catalogo, ensure_ascii=False, indent=1), encoding="utf-8")
    C.reemplazar_atomico(tmp, C.RUTA_CATALOGO)


def guardada(ficha: dict, nota: str) -> dict:
    """Última posición guardada del buque (puede ser vieja: la plantilla muestra su antigüedad)."""
    if ficha.get("ais"):
        return {"pos": ficha["ais"], "origen": "guardada", "consultado": ficha.get("ais_consultado"), "nota": nota}
    return {"pos": None, "nota": nota}


def posiciones(cola: list, catalogo: dict, ahora) -> tuple:
    """nombre -> {'pos', 'origen', 'consultado', 'nota'}. Caché primero; la web solo para lo vencido o faltante."""
    res = {}
    resumen = {"activo": C.USAR_AIS, "en_cola": len(cola), "consultados": 0, "de_cache": 0, "frenado": None}
    if not cola:
        return res, resumen
    if not C.USAR_AIS:
        return {b["nombre"]: guardada(catalogo.get(b["nombre"], {}), "consulta web apagada (USAR_AIS = False)")
                for b in cola}, resumen
    try:
        sesion = abrir_sesion()
    except ImportError:
        log_radar.warning("falta curl_cffi (pip install curl_cffi): el radar sale con las posiciones guardadas")
        resumen["frenado"] = "falta curl_cffi"
        return {b["nombre"]: guardada(catalogo.get(b["nombre"], {}), "falta curl_cffi") for b in cola}, resumen
    web = Web(sesion, PAUSA_S, MAX_CONSULTAS, CODIGOS_FRENO, TIMEOUT_S)
    sello = ahora.strftime("%Y-%m-%d %H:%M")
    try:
        for b in cola:
            nombre = b["nombre"]
            ficha = catalogo.setdefault(nombre, {})
            if ficha.get("ais") and C.edad_horas(ficha.get("ais_consultado"), ahora) < getattr(C, "HORAS_CACHE_AIS_VIGENTE", HORAS_CACHE_AIS):
                res[nombre] = guardada(ficha, None) | {"origen": "caché"}
                resumen["de_cache"] += 1
                continue
            if web.frenado:
                res[nombre] = guardada(ficha, f"sin consulta: {web.frenado}")
                continue
            if not (ficha.get("imo") and ficha.get("mmsi")):
                if C.edad_horas(ficha.get("no_encontrado"), ahora) < DIAS_REINTENTO * 24:
                    res[nombre] = guardada(ficha, f"ProveedorA no lo encontró el {ficha['no_encontrado'][:10]}; "
                                                  f"se reintenta a los {DIAS_REINTENTO} días")
                    continue
                ids = buscar_imo_mmsi(web, nombre)
                if "error" in ids:
                    if ids["error"] == "ProveedorA no encontró el buque":
                        ficha["no_encontrado"] = sello
                    res[nombre] = guardada(ficha, ids["error"])
                    continue
                ficha.pop("no_encontrado", None)
                ficha.update(imo=ids["imo"], mmsi=ids["mmsi"], nombre_web=ids["nombre_web"], exacto=ids["exacto"])
            pos = posicion_proveedor_b(web, nombre, ficha["imo"], ficha["mmsi"], V.CALLAO["locode"])
            resumen["consultados"] += 1
            if "error" in pos:
                res[nombre] = guardada(ficha, pos["error"])
                continue
            pos["dist_km"] = distancia_km(pos["lat"], pos["lon"], V.CALLAO["lat"], V.CALLAO["lon"])
            ficha.update(ais=pos, ais_consultado=sello)
            res[nombre] = {"pos": pos, "origen": "consulta", "consultado": sello, "nota": None}
    finally:
        resumen["frenado"] = web.frenado
        guardar_catalogo(catalogo)
        getattr(sesion, "close", lambda: None)()
    return res, resumen


def obtener_mapa():
    """Países del mundo (world-atlas), bajados una vez y guardados junto al script. Sin mapa el radar sale igual."""
    if C.RUTA_MAPA.exists():
        try:
            return json.loads(C.RUTA_MAPA.read_text(encoding="utf-8"))
        except ValueError:
            log_radar.warning("%s ilegible: se vuelve a bajar", C.RUTA_MAPA.name)
    for url in URLS_MAPA:
        try:
            with urllib.request.urlopen(url, timeout=30) as resp:
                mapa = json.loads(resp.read().decode("utf-8"))
            C.RUTA_MAPA.write_text(json.dumps(mapa, separators=(",", ":")), encoding="utf-8")
            log_radar.info("mapa del mundo bajado de %s (una sola vez)", url)
            return mapa
        except (OSError, ValueError) as e:
            log_radar.warning("no se pudo bajar el mapa de %s (%s)", url, e)
    return None


def escribir_html(datos: dict, mapa) -> None:
    plantilla = C.RUTA_PLANTILLA.read_text(encoding="utf-8")
    for marca in (MARCA_DATOS, MARCA_MAPA):
        assert plantilla.count(marca) == 1, f"{C.RUTA_PLANTILLA.name} sin la marca {marca}: ¿plantilla equivocada?"
    a_js = lambda obj: json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html_out = plantilla.replace(MARCA_MAPA, a_js(mapa)).replace(MARCA_DATOS, a_js(datos))
    C.RUTA_SALIDA_HTML.parent.mkdir(parents=True, exist_ok=True)
    tmp = C.RUTA_SALIDA_HTML.with_name(C.RUTA_SALIDA_HTML.stem + ".tmp.html")
    tmp.write_text(html_out, encoding="utf-8")
    C.reemplazar_atomico(tmp, C.RUTA_SALIDA_HTML)


def main_radar(armado=None) -> None:
    t0 = time.perf_counter()
    ahora = pd.Timestamp.now().floor("s")
    hoy = ahora.normalize()
    log_radar.info("=== RADAR · llegadas al Callao ===")
    assert C.RUTA_PLANTILLA.exists(), f"Falta {C.RUTA_PLANTILLA.name} junto al script"
    mar, fuera, corte = V.leer_maritimos()
    log_radar.info("seguimiento: %d líneas marítimas en el radar (%d BL) · fuera: %d aéreas, %d marítimas sin BL, "
                   "%d entregadas al DC", len(mar), mar["bl"].nunique(),
                   fuera["aereas"], fuera["sin_bl"], fuera["entregadas"])
    largo, aviso_pedidos = V.lineas_con_pedidos(mar)
    sold = C.sold_to_por_pedido(avisar_sin_columna=False)
    emb = C.leer_embarques(C.COLS_EMBARQUE_RADAR, deduplicar=True, logger=log_radar)
    buques = V.agrupar_buques(V.armar_bls(largo, sold, emb, hoy))
    sin_nombre = sum(1 for b in buques if not b["nombre"])
    if sin_nombre:
        log_radar.warning("%d BL sin name_of_ocean_vessel en el reporte: se listan sin posición", sin_nombre)

    orden = {"vencida": 0, "sin_fecha": 1, "por_arribar": 2}
    cola = sorted([b for b in buques if b["nombre"] and b["por_fecha"] in orden],
                  key=lambda b: (orden[b["por_fecha"]], b["dias_hasta"] if b["dias_hasta"] is not None else 0))
    log_radar.info("cola del disparador: %d buques · %s", len(cola), ", ".join(b["nombre"] for b in cola) or "ninguno")
    catalogo = leer_catalogo()
    ais, resumen = posiciones(cola, catalogo, ahora)
    for b in buques:
        V.aplicar_estado(b, ais.get(b["nombre"]) if b["nombre"] in ais else None,
                         catalogo.get(b["nombre"] or "", {}), ahora)
    log_radar.info("AIS: %d consultados en la web · %d de caché · %s", resumen["consultados"], resumen["de_cache"],
                   f"FRENADO ({resumen['frenado']})" if resumen["frenado"] else "sin freno")
    for b in cola:
        log_radar.info("   %-24s %s", b["nombre"], f"{b['ais']['dist_km']:,} km del Callao ({b['ais_origen']})"
                       if b["ais"] else f"sin posición: {b['ais_nota']}")
    vista_mar, fuera_vista = V.separar_vista(buques)
    estados = pd.Series([b["estado"] for b in vista_mar], dtype="string").value_counts().to_dict()
    log_radar.info("pestaña marítima: %d buques con nombre y cerca %s · fuera de la vista: %d más lejos, %d sin fecha ni "
                   "posición, %d BL sin nombre de buque, %d BL ya arribados · %d avisos anticipados", len(vista_mar),
                   estados, fuera_vista["lejos"], fuera_vista["sin_fecha"], fuera_vista["sin_nombre_bl"],
                   fuera_vista["arribados_bl"], fuera_vista["avisos_anticipados"])

    aereo = V.aereos_en_terminal(armado=armado)
    hawbs = aereo["hawbs"]
    log_radar.info("pestaña aérea: %s", f"SIN DATOS ({aereo['error']})" if aereo["error"] else
                   f"{len(hawbs)} HAWB con aviso y sin salida a DC · {sum(h['fuera_sla'] for h in hawbs)} fuera de SLA "
                   f"({aereo['sla_hab']} d.h.) · {sum(h['alerta_aviso'] for h in hawbs)} con bultos malos o faltantes · "
                   f"{aereo['en_ruta']} con salida a DC programada (no entran)")

    datos = {"generado": ahora.strftime("%Y-%m-%d %H:%M"), "corte_bi": pd.Timestamp(corte).strftime("%Y-%m-%d %H:%M"),
             "callao": V.CALLAO, "radio_cerca_km": V.RADIO_CERCA_KM, "dias_proximo": V.DIAS_PROXIMO,
             "link_detalle": LINK_DETALLE or C.RUTA_IN_TRANSIT.as_uri(), "nombre_detalle": C.RUTA_IN_TRANSIT.name,
             "aviso_pedidos": aviso_pedidos, "ais": resumen,
             "maritimo": {"buques": vista_mar, "fuera": fuera | fuera_vista}, "aereo": aereo,
             "alertas_eta": V.alertas_eta(), "salida_dc": V.salida_dc()}   # paso 5
    escribir_html(datos, obtener_mapa())
    log_radar.info("radar: %d buques · %d HAWB · en la carta %d -> %s", len(vista_mar), len(hawbs),
                   sum(1 for b in vista_mar if b["ais"]), C.RUTA_SALIDA_HTML)
    log_radar.info("=== FIN (%.0f s) ===", time.perf_counter() - t0)
    if C.ABRIR_NAVEGADOR:
        webbrowser.open(C.RUTA_SALIDA_HTML.as_uri())


def main_salidas() -> None:
    """Ejecuta correo y radar en el mismo proceso compartiendo vista.armar() en memoria (R1)."""
    armado = main_correo()
    main_radar(armado=armado)
