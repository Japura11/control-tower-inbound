"""Demo del radar con VARIOS buques (sintéticos) y el mapa mundial real, para juzgar la Carta.
Parte de los datos del día 2 de salida_paso4 y agrega buques en todos los estados, en rutas reales del Pacífico.
Uso: python demo_mapa/demo_varios.py [plantilla]  -> demo_mapa/demo_varios.html
"""
import copy
import json
import math
import re
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parent
PLANTILLA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "torre" / "radar_embarques_plantilla.html"
CALLAO = (-12.0464, -77.1428)

html = (RAIZ / "salida_demo" / "dia2" / "radar_embarques.html").read_text(encoding="utf-8")
D = json.loads(re.search(r"const D = (\{.*?\});\s*\n", html, re.S).group(1))
base = next(b for b in D["maritimo"]["buques"] if b["nombre"] == "SINT AURORA")


def km(lat, lon):
    r = math.radians
    a = math.sin(r(lat - CALLAO[0]) / 2) ** 2 + math.cos(r(lat)) * math.cos(r(CALLAO[0])) * math.sin(r(lon - CALLAO[1]) / 2) ** 2
    return int(6371 * 2 * math.asin(math.sqrt(a)))


# nombre, estado, lat, lon (None = sin AIS), días desde/hasta, BL, unidades, 433
NUEVOS = [
    ("SINT ANDES", "por_arribar", -5.2, -81.9, 1, 2, 420),
    ("SINT GALAPAGOS", "por_arribar", -1.5, -83.0, 2, 3, 260),
    ("SINT PANAMA STAR", "por_arribar", 3.8, -80.4, 3, 3, 610),
    ("SINT LIMA BAY", "vencida", -11.4, -77.9, -2, -1, 150),
    ("SINT PISCO", "vencida", -13.9, -76.9, -4, -3, 95),
    ("SINT COSTA AZUL", "en_callao", -12.04, -77.16, 0, 0, 330),
    ("SINT TAIPEI", "por_arribar", -8.9, -79.4, 1, 3, 780),
    ("SINT SHANGHAI EXP", "por_arribar", -9.8, -80.6, 2, 3, 1200),
    ("SINT CHANCAY", "en_puerto", None, None, None, None, 240),
    ("SINT MOLLENDO", "en_puerto", None, None, None, None, 60),
]
for i, (nombre, estado, lat, lon, d, h, u) in enumerate(NUEVOS):
    b = copy.deepcopy(base)
    b.update(id=f"d{i}", nombre=nombre, estado=estado, por_fecha=estado, dias_desde=d, dias_hasta=h, dias_hasta_max=h,
             imo=f"99900{i:02d}", mmsi=f"7770001{i:02d}")
    if lat is None:
        b.update(ais=None, ais_valida=False, ais_origen=None)
    else:
        b["ais"] = {**b["ais"], "lat": lat, "lon": lon, "dist_km": km(lat, lon), "arribado": estado == "en_callao"}
    for x in b["bls"]:
        x.update(bl=f"SINTM-3{i:05d}", buque=nombre, unidades=float(u),
                 hito="Arribado" if estado == "en_puerto" else "En tránsito internacional")
        for ln in x["lineas"]:
            ln.update(so=f"42298{i:05d}", cantidad=float(u))
            for p in ln["pedidos"]:
                p.update(so433=f"43398{i:05d}")
    D["maritimo"]["buques"].append(b)
D["ais"] = {**D["ais"], "en_cola": len(NUEVOS) + 1, "consultados": len(NUEVOS) + 1}

plantilla = PLANTILLA.read_text(encoding="utf-8")
mapa = (RAIZ / "assets" / "countries-50m.json").read_text(encoding="utf-8")
salida = plantilla.replace("/*__MAPA__*/null", mapa).replace("/*__DATOS__*/{}", json.dumps(D, ensure_ascii=False))
destino = RAIZ / "salida_demo" / "demo_varios.html"
destino.write_text(salida, encoding="utf-8")
print(destino, len(D["maritimo"]["buques"]), "buques")
