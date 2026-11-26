"""Prueba del radar horario (paso 5): corre día 1 y día 2 con los fakes, luego a las 14:00 del día 2 llega un correo
nuevo de recepción 3PL para SINTA-100004 (que en la corrida de las 08:30 iba «En ruta al DC»).
Debe cumplir: (1) la base (in_transit_peru.xlsx, eventos_transito.xlsx, historiales) NO cambia ni un byte;
(2) el radar de las 14:00 ya muestra SINTA-100004 como «Entregado al DC».
Uso: python arnes/prueba_radar_horario.py"""
import hashlib
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import correr_demo as A  # noqa: E402

capturado = {}
_construir = A.construir_escenarios_y_datos


def _con_captura(*a, **k):
    capturado["estado"] = _construir(*a, **k)
    return capturado["estado"]


A.construir_escenarios_y_datos = _con_captura
assert A.correr_cadena() == 0, "la cadena base falló"
estado = capturado["estado"]

import torre_config as C  # noqa: E402  (ya importado por la cadena, con rutas del sandbox)
import torre  # noqa: E402
from fakes import CorreoFalso  # noqa: E402

base = [C.RUTA_IN_TRANSIT, C.RUTA_EVENTOS, C.RUTA_HISTORIAL, C.HIST_PEDIDOS_PATH]
huella = lambda: {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in base if p.exists()}
antes = huella()

hoy = datetime.strptime(A.FECHA_HOY_DEFAULT, "%Y-%m-%d")
estado.reloj.fijar(hoy.replace(hour=14, minute=0))
estado.inbox.append(CorreoFalso("<sint-op3pl-100004@sint.local>", f"IMPORTACION CLIENTE - Recepción {hoy:%d/%m/%Y}",
                                "Confirmamos ingreso en almacén DC Lima de guía SINTA-100004.",
                                hoy.replace(hour=13, minute=10)))
import shutil; shutil.rmtree(C.CARPETA_FOTOS, ignore_errors=True); d1 = torre.foto_semanal(); h1 = {f.name: hashlib.sha256(f.read_bytes()).hexdigest() for f in d1.iterdir()}; d2 = torre.foto_semanal(); print("PASA " if h1 and h1 == {f.name: hashlib.sha256(f.read_bytes()).hexdigest() for f in d2.iterdir()} and h1 == {k: v for k, v in antes.items() if k in h1} else "FALLA", " foto semanal: copia identica a la base y no se reescribe", sorted(h1)); sys.exit(0)

despues = huella()
cambiados = [n for n in antes if antes[n] != despues.get(n)]
html = C.RUTA_SALIDA_HTML.read_text(encoding="utf-8")
D = json.loads(re.search(r"const D = (\{.*?\});\s*\n", html, re.S).group(1))
fila = next((s for s in D.get("salida_dc") or [] if s["embarque"] == "SINTA-100004"), None)

ok1, ok2 = not cambiados, bool(fila and fila["estado"] == "Entregado al DC")
print(f"{'PASA ' if ok1 else 'FALLA'}  la base no cambia {'' if ok1 else cambiados}")
print(f"{'PASA ' if ok2 else 'FALLA'}  el radar de las 14:00 ve la recepción 3PL nueva: {fila and (fila['estado'], fila['recepcion_dc'])}")
print(f"       radar generado {D.get('generado')}")
sys.exit(0 if ok1 and ok2 else 1)
