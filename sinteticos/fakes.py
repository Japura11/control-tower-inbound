# -*- coding: utf-8 -*-
"""
fakes.py — Módulos falsos (win32com, pythoncom, pdfplumber, lake_core, curl_cffi) y reloj virtual
para ejecutar la cadena actual en memoria/proceso sin tocar Outlook, Excel COM, Lake ni red real.
"""

from __future__ import annotations

import base64
import json
import os
import re
import shutil
import sys
import types
import urllib.request
import webbrowser
from datetime import datetime as _RealDatetime, timedelta
from pathlib import Path

import pandas as pd

_ORIG_TIMESTAMP = pd.Timestamp
_MARCA_PDF = b"%SINT_PAGES_B64:"


def escribir_pdf_sintetico(ruta: Path, paginas: list[str]) -> Path:
    """Escribe un archivo .pdf mínimo válido con texto sintético codificado (o vacío si simula escaneado)."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    b64 = base64.b64encode(json.dumps(paginas, ensure_ascii=False).encode("utf-8"))
    contenido = (
        b"%PDF-1.4\r\n"
        + _MARCA_PDF + b64 + b"\r\n"
        + b"1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\r\n"
        + b"2 0 obj << /Type /Pages /Kids [] /Count 0 >> endobj\r\n"
        + b"%%EOF\r\n"
    )
    ruta.write_bytes(contenido)
    return ruta


def leer_paginas_pdf_sintetico(ruta: Path) -> list[str]:
    datos = Path(ruta).read_bytes()
    for linea in datos.splitlines():
        if linea.startswith(_MARCA_PDF):
            crudo = base64.b64decode(linea[len(_MARCA_PDF):].strip())
            return json.loads(crudo.decode("utf-8"))
    return [""]


# ══════════════════════════ RELOJ VIRTUAL DETERMINISTA ══════════════════════
class RelojVirtual:
    def __init__(self, ahora: _RealDatetime):
        self.ahora: _RealDatetime = ahora

    def fijar(self, ahora: _RealDatetime) -> None:
        self.ahora = ahora


def crear_proxies_reloj(reloj: RelojVirtual):
    """Crea proxies de datetime y pandas que interceptan únicamente .now() conservando tipos reales."""

    class _MetaDatetime(type):
        def __instancecheck__(cls, instance):
            return isinstance(instance, _RealDatetime)

        def __subclasscheck__(cls, subclass):
            return issubclass(subclass, _RealDatetime)

    class FakeDatetime(metaclass=_MetaDatetime):
        min = _RealDatetime.min
        max = _RealDatetime.max
        resolution = _RealDatetime.resolution

        def __new__(cls, *args, **kwargs):
            return _RealDatetime(*args, **kwargs)

        @classmethod
        def now(cls, tz=None):
            if tz is not None:
                return _RealDatetime.fromtimestamp(reloj.ahora.timestamp(), tz=tz)
            return reloj.ahora

        @classmethod
        def today(cls):
            return reloj.ahora

        @classmethod
        def fromtimestamp(cls, ts, tz=None):
            return _RealDatetime.fromtimestamp(ts, tz=tz)

        @classmethod
        def fromisoformat(cls, s):
            return _RealDatetime.fromisoformat(s)

        @classmethod
        def strptime(cls, date_string, fmt):
            return _RealDatetime.strptime(date_string, fmt)

        @classmethod
        def combine(cls, date, time, tzinfo=True):
            return _RealDatetime.combine(date, time) if tzinfo is True else _RealDatetime.combine(date, time, tzinfo)

    class _MetaTimestamp(type):
        def __instancecheck__(cls, instance):
            return isinstance(instance, _ORIG_TIMESTAMP)

        def __subclasscheck__(cls, subclass):
            return issubclass(subclass, _ORIG_TIMESTAMP)

    class FakeTimestamp(metaclass=_MetaTimestamp):
        min = _ORIG_TIMESTAMP.min
        max = _ORIG_TIMESTAMP.max

        def __new__(cls, *args, **kwargs):
            return _ORIG_TIMESTAMP(*args, **kwargs)

        @classmethod
        def now(cls, tz=None):
            base = _ORIG_TIMESTAMP(reloj.ahora)
            return base.tz_localize(tz) if tz is not None else base

        @classmethod
        def today(cls, tz=None):
            return cls.now(tz=tz)

        @classmethod
        def fromtimestamp(cls, ts, tz=None):
            return _ORIG_TIMESTAMP.fromtimestamp(ts, tz=tz)

    class PandasProxy:
        def __init__(self, modulo_pd):
            self._pd = modulo_pd
            self.Timestamp = FakeTimestamp

        def __getattr__(self, nombre):
            return getattr(self._pd, nombre)

    return FakeDatetime, PandasProxy(pd)


def aplicar_reloj_a_modulos(modulos: list, reloj: RelojVirtual) -> None:
    fake_dt, fake_pd = crear_proxies_reloj(reloj)
    for mod in modulos:
        if hasattr(mod, "datetime"):
            setattr(mod, "datetime", fake_dt)
        if hasattr(mod, "pd"):
            setattr(mod, "pd", fake_pd)


# ══════════════════════════ ESTADO DEL ENTORNO FALSO ════════════════════════
class AdjuntoMail:
    def __init__(self, ruta: Path, nombre: str | None = None):
        self._ruta = Path(ruta)
        self.FileName = nombre or self._ruta.name

    def SaveAsFile(self, destino: str) -> None:
        shutil.copy2(self._ruta, destino)


class CorreoFalso:
    Class = 43  # olMail

    def __init__(
        self,
        id_internet: str,
        asunto: str,
        cuerpo: str,
        fecha: _RealDatetime,
        html_body: str = "",
        adjuntos: list[AdjuntoMail] | None = None,
    ):
        self.EntryID = f"ENTRY-{re.sub(r'[^A-Za-z0-9]', '', id_internet)}"
        self.Subject = asunto
        self.Body = cuerpo
        self.HTMLBody = html_body
        self.ReceivedTime = fecha
        self.SentOn = fecha
        self.Attachments = list(adjuntos or [])
        self.PropertyAccessor = types.SimpleNamespace(
            GetProperty=lambda _prop, mid=id_internet: mid
        )


class ItemsOutlook(list):
    def Restrict(self, filtro: str):
        m = re.match(r"^\[(\w+)\]\s*>=\s*'([^']+)'$", filtro.strip())
        if not m:
            raise ValueError(f"Filtro Outlook no soportado en fake: {filtro!r}")
        campo, texto_fecha = m.group(1), m.group(2)
        desde = _RealDatetime.strptime(texto_fecha, "%m/%d/%Y %I:%M %p")
        filtrados = [
            item for item in self
            if getattr(item, campo).replace(second=0, microsecond=0) >= desde
        ]
        return ItemsOutlook(filtrados)

    def Sort(self, propiedad: str = "[ReceivedTime]", descendente: bool = False) -> None:
        campo = propiedad.strip("[]")
        self.sort(key=lambda item: getattr(item, campo), reverse=bool(descendente))


class EstadoFakes:
    """Contenedor central de datos sintéticos y capturas para una corrida de Día 1 / Día 2."""

    def __init__(
        self,
        reloj: RelojVirtual,
        om_por_dia: dict[int, Path],
        hus_por_dia: dict[int, pd.DataFrame],
        inbox: list[CorreoFalso],
        shipments: list[CorreoFalso],
        enviados: list[CorreoFalso],
        barcos_ais: dict[str, dict],
        carpeta_captura_mail: Path,
    ):
        self.reloj = reloj
        self.dia_activo = 1
        self.om_por_dia = om_por_dia
        self.hus_por_dia = hus_por_dia
        self.inbox = inbox
        self.shipments = shipments
        self.enviados = enviados
        self.barcos_ais = barcos_ais
        self.carpeta_captura_mail = carpeta_captura_mail

        self.logs_lake: list[str] = []
        self.correos_emitidos: list[dict] = []
        self.llamadas_ais: list[str] = []
        self.lecturas_pdf: list[str] = []
        self.descargas_adjuntos_enviados: list[str] = []

    def activar_dia(self, dia: int, momento: _RealDatetime, carpeta_captura: Path) -> None:
        self.dia_activo = dia
        self.reloj.fijar(momento)
        self.carpeta_captura_mail = carpeta_captura
        self.carpeta_captura_mail.mkdir(parents=True, exist_ok=True)


# ══════════════════════════ INSTALACIÓN EN SYS.MODULES ══════════════════════
def instalar_fakes(estado: EstadoFakes) -> None:
    """Inyecta módulos falsos en sys.modules ANTES de importar cualquier script de base_actual."""

    # 1. pythoncom
    mod_pythoncom = types.ModuleType("pythoncom")
    mod_pythoncom.CoInitialize = lambda: None
    mod_pythoncom.CoUninitialize = lambda: None
    sys.modules["pythoncom"] = mod_pythoncom

    # 2. lake_core
    mod_lc = types.ModuleType("lake_core")

    def lc_log(mensaje: str) -> None:
        linea = f"[{estado.reloj.ahora:%H:%M:%S}] {mensaje}"
        estado.logs_lake.append(linea)
        print(linea)

    def lc_get_connection():
        return types.SimpleNamespace(closed=False)

    def lc_query_chunked(conn, sql: str, keys: list, chunk_size: int = 1000) -> pd.DataFrame:
        hus = estado.hus_por_dia[estado.dia_activo]
        conjunto = {str(k).strip() for k in keys}
        sub = hus[hus["delivery_number"].astype(str).isin(conjunto)].copy()
        return sub.reset_index(drop=True)

    def lc_atomic_to_excel(hojas: dict, ruta: Path, meta_texto: str = "", _extra=None) -> None:
        destino = Path(ruta)
        destino.parent.mkdir(parents=True, exist_ok=True)
        tmp = destino.with_name(f"{destino.stem}.tmp.xlsx")
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            for nombre_hoja, df in hojas.items():
                df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            if meta_texto:
                pd.DataFrame([{"generado_en": estado.reloj.ahora.strftime("%Y-%m-%d %H:%M:%S"),
                               "detalle": str(meta_texto)}]).to_excel(writer, sheet_name="meta", index=False)
        os.replace(tmp, destino)
        ts = estado.reloj.ahora.timestamp()
        os.utime(destino, (ts, ts))
        lc_log(f"    atomic_to_excel -> {destino.name} ({', '.join(hojas.keys())})")

    mod_lc.log = lc_log
    mod_lc.get_connection = lc_get_connection
    mod_lc.query_chunked = lc_query_chunked
    mod_lc.atomic_to_excel = lc_atomic_to_excel
    sys.modules["lake_core"] = mod_lc

    # 3. pdfplumber
    mod_pdf = types.ModuleType("pdfplumber")

    class _FakePage:
        def __init__(self, texto: str):
            self._texto = texto

        def extract_text(self) -> str:
            return self._texto

    class _FakePDF:
        def __init__(self, ruta: str | Path):
            paginas = leer_paginas_pdf_sintetico(Path(ruta))
            self.pages = [_FakePage(t) for t in paginas]

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

    mod_pdf.open = lambda ruta: _FakePDF(ruta)
    sys.modules["pdfplumber"] = mod_pdf

    # 4. win32com & win32com.client (Excel + Outlook)
    class _FakeConnections(list):
        @property
        def Count(self) -> int:
            return len(self)

    class _FakeWorkbook:
        def __init__(self, ruta: Path):
            self.ruta = Path(ruta)
            self.ReadOnly = False
            self._refrescado = False
            self.Connections = _FakeConnections([
                types.SimpleNamespace(OLEDBConnection=types.SimpleNamespace(BackgroundQuery=True)),
                types.SimpleNamespace(ODBCConnection=types.SimpleNamespace(BackgroundQuery=True)),
            ])

        def RefreshAll(self) -> None:
            self._refrescado = True

        def Save(self) -> None:
            mtime_previo = self.ruta.stat().st_mtime if self.ruta.exists() else 0.0
            if self._refrescado and estado.dia_activo in estado.om_por_dia:
                fuente = estado.om_por_dia[estado.dia_activo]
                tmp = self.ruta.with_name(f"{self.ruta.stem}.refresh.tmp.xlsx")
                shutil.copy2(fuente, tmp)
                os.replace(tmp, self.ruta)
            nuevo_ts = max(estado.reloj.ahora.timestamp(), mtime_previo + 1.0)
            os.utime(self.ruta, (nuevo_ts, nuevo_ts))

        def Close(self, SaveChanges: bool = False) -> None:
            pass

    class _FakeExcelApp:
        Visible = False
        DisplayAlerts = False

        def __init__(self):
            self.Workbooks = types.SimpleNamespace(
                Open=lambda ruta, UpdateLinks=0, ReadOnly=False: _FakeWorkbook(Path(ruta))
            )

        def CalculateUntilAsyncQueriesDone(self) -> None:
            pass

        def Quit(self) -> None:
            pass

    class _AdjuntoRastreado:
        def __init__(self, orig: AdjuntoMail, es_enviado: bool = False):
            self._orig = orig
            self.FileName = orig.FileName
            self._es_enviado = es_enviado

        def SaveAsFile(self, destino: str) -> None:
            if self._es_enviado:
                estado.descargas_adjuntos_enviados.append(self.FileName)
            elif self.FileName.lower().endswith(".pdf"):
                estado.lecturas_pdf.append(self.FileName)
            self._orig.SaveAsFile(destino)

    def _clonar_visibles(lista_correos: list[CorreoFalso], attr_fecha: str, es_enviado: bool = False) -> ItemsOutlook:
        tope = estado.reloj.ahora
        visibles = []
        for c in lista_correos:
            if getattr(c, attr_fecha) <= tope:
                copia = CorreoFalso(
                    id_internet=c.PropertyAccessor.GetProperty(""),
                    asunto=c.Subject,
                    cuerpo=c.Body,
                    fecha=c.ReceivedTime,
                    html_body=c.HTMLBody,
                    adjuntos=[_AdjuntoRastreado(a, es_enviado=es_enviado) for a in c.Attachments],
                )
                copia.SentOn = c.SentOn
                copia.EntryID = c.EntryID
                visibles.append(copia)
        return ItemsOutlook(visibles)

    class _FakeStore:
        DisplayName = "operaciones@cliente.com"

        def GetDefaultFolder(self, codigo: int):
            if codigo == 6:
                return types.SimpleNamespace(
                    Name="Inbox",
                    Items=_clonar_visibles(estado.inbox, "ReceivedTime"),
                )
            raise KeyError(f"DefaultFolder no soportado en Store: {codigo}")

        def GetRootFolder(self):
            return types.SimpleNamespace(
                Folders={
                    "Shipments": types.SimpleNamespace(
                        Name="Shipments",
                        Items=_clonar_visibles(estado.shipments, "ReceivedTime"),
                    )
                }
            )

    class _FakeDraftMail:
        def __init__(self):
            self.To = ""
            self.CC = ""
            self.Subject = ""
            self.HTMLBody = ""
            self.adjuntos_copiados: list[Path] = []
            self.Attachments = types.SimpleNamespace(Add=self._agregar_adjunto)

        def _agregar_adjunto(self, ruta_str: str) -> None:
            origen = Path(ruta_str)
            destino = estado.carpeta_captura_mail / "mercaderia_arribada.xlsx"
            destino.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(origen, destino)
            self.adjuntos_copiados.append(destino)

        def _guardar_borrador(self, modo: str) -> None:
            destino_html = estado.carpeta_captura_mail / "correo_arribados.html"
            destino_html.parent.mkdir(parents=True, exist_ok=True)
            documento = (
                '<!DOCTYPE html>\n<html><head><meta charset="utf-8">'
                f"<title>{self.Subject}</title></head>"
                '<body style="margin:0;padding:0;background:#E6E2E4">'
                f"{self.HTMLBody}</body></html>"
            )
            destino_html.write_text(documento, encoding="utf-8")
            estado.correos_emitidos.append({
                "dia": estado.dia_activo,
                "modo": modo,
                "asunto": self.Subject,
                "para": self.To,
                "cc": self.CC,
                "html_path": str(destino_html),
                "adjuntos": [str(p) for p in self.adjuntos_copiados],
            })

        def Display(self, _modal: bool = False) -> None:
            # DRY_RUN sin abrir ventana: solo persiste el HTML en disco
            self._guardar_borrador("DRY_RUN_DISPLAY")

        def Send(self) -> None:
            self._guardar_borrador("SEND")

    class _FakeOutlookApp:
        def GetNamespace(self, nombre: str):
            assert nombre == "MAPI"
            return types.SimpleNamespace(
                Stores=[_FakeStore()],
                GetDefaultFolder=lambda codigo: types.SimpleNamespace(
                    Name="Elementos enviados",
                    Items=_clonar_visibles(estado.enviados, "SentOn", es_enviado=True),
                ) if codigo == 5 else (_ for _ in ()).throw(KeyError(codigo)),
            )

        def CreateItem(self, tipo: int = 0):
            return _FakeDraftMail()

    def _dispatch(prog_id: str):
        nombre = prog_id.lower()
        if "excel" in nombre:
            return _FakeExcelApp()
        if "outlook" in nombre:
            return _FakeOutlookApp()
        raise ValueError(f"COM ProgID inesperado en fake: {prog_id!r}")

    mod_win32 = types.ModuleType("win32com")
    mod_win32_client = types.ModuleType("win32com.client")
    mod_win32_client.Dispatch = _dispatch
    mod_win32_client.DispatchEx = _dispatch
    mod_win32.client = mod_win32_client
    sys.modules["win32com"] = mod_win32
    sys.modules["win32com.client"] = mod_win32_client

    # 5. curl_cffi & curl_cffi.requests (AIS sin red)
    class _FakeResp:
        def __init__(self, status_code: int, text: str):
            self.status_code = status_code
            self.text = text

    def _html_proveedor_b(ficha: dict) -> str:
        payload = {
            "vessel": {
                "name": ficha["nombre"],
                "imo": ficha["imo"],
                "mmsi": ficha["mmsi"],
                "position": {"latitude": ficha["lat"], "longitude": ficha["lon"]},
                "destinationUnlocode": ficha["destino"],
                "eta": ficha["eta_iso"],
                "lastUpdatedAt": ficha["actualizado_iso"],
                "flag": ficha.get("bandera", "PA"),
                "teu": ficha.get("teu", 12000),
                "draught": ficha.get("calado", 11.8),
            },
            "futurePortCalls": {
                "current_port": ficha["escalas_pasadas"][0]["portUnlocode"] if ficha["escalas_pasadas"] else "PABLB",
                "next_port": ficha["escalas_futuras"][0]["port_unlocode"] if ficha["escalas_futuras"] else "PECLL",
                "future_port_calls": ficha["escalas_futuras"],
            },
            "vesselPortCalls": {"data": ficha["escalas_pasadas"]},
            "ports": [{"unlocode": "PECLL", "latitude": -12.0464, "longitude": -77.1428}],
        }
        escapado = json.dumps(payload).replace('"', '\\"')
        return f'<html><body><script>self.__next_f.push([1,"{escapado}"])</script></body></html>'

    class _FakeCffiSession:
        def __init__(self, impersonate: str | None = None):
            self.impersonate = impersonate

        def close(self) -> None:
            pass

        def get(self, url: str, timeout: int = 20) -> _FakeResp:
            estado.llamadas_ais.append(url)
            barcos_dia = estado.barcos_ais.get(estado.dia_activo, {})
            if "ais-proveedor-a.example/vessels?name=" in url:
                nombre = urllib.request.unquote(url.split("name=", 1)[1]).upper().strip()
                if nombre not in barcos_dia:
                    return _FakeResp(200, "<html><body><p>No results</p></body></html>")
                imo = barcos_dia[nombre]["imo"]
                return _FakeResp(
                    200,
                    f'<html><body><a class="ship-link" href="/vessels/details/{imo}">'
                    f'<div class="slna">{nombre}</div></a></body></html>',
                )
            if "ais-proveedor-a.example/vessels/details/" in url:
                imo = url.rsplit("/", 1)[1]
                ficha = next((v for v in barcos_dia.values() if v["imo"] == imo), None)
                if not ficha:
                    return _FakeResp(404, "Not found")
                return _FakeResp(
                    200,
                    f'<html><body><div id="djson" data-json=\'{{"mmsi": "{ficha["mmsi"]}"}}\'></div></body></html>',
                )
            if "ais-proveedor-b.example/vessel/" in url:
                ficha = next((v for v in barcos_dia.values() if f"imo:{v['imo']}" in url), None)
                if not ficha:
                    return _FakeResp(404, "Not found")
                return _FakeResp(200, _html_proveedor_b(ficha))
            return _FakeResp(404, "Not mocked")

    mod_cffi = types.ModuleType("curl_cffi")
    mod_cffi_req = types.ModuleType("curl_cffi.requests")
    mod_cffi_req.Session = _FakeCffiSession
    mod_cffi.requests = mod_cffi_req
    sys.modules["curl_cffi"] = mod_cffi
    sys.modules["curl_cffi.requests"] = mod_cffi_req

    # Bloqueo explícito de cualquier red accidental por urllib o apertura de navegador
    def _bloquear_urlopen(*args, **kwargs):
        raise OSError("Llamada de red bloqueada por el arnés sintético")

    urllib.request.urlopen = _bloquear_urlopen
    webbrowser.open = lambda *a, **kw: False
