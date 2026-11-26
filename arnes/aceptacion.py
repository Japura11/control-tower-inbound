"""Pruebas de aceptación del paso 5: lo que pidió el usuario el 25/11, medido sobre el radar del día 2.

Lee el JSON que el radar trae embebido (const D = {...}) y revisa cada escenario de sinteticos/ESCENARIOS.md.
Hoy (v2, paso 3) deben FALLAR: documentan las brechas. El paso 5 está listo cuando pasan todas.

Uso: python arnes/aceptacion_paso5.py [carpeta_salida]   (por defecto salida_v2/dia2)
"""
import json
import re
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
CARPETA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "salida_demo" / "dia2"


def datos_radar(carpeta: Path) -> dict:
    html = (carpeta / "radar_embarques.html").read_text(encoding="utf-8")
    m = re.search(r"const D = (\{.*?\});\s*\n", html, re.S)
    assert m, "no encuentro 'const D = {...};' en el radar"
    return json.loads(m.group(1))


def texto(obj) -> str:
    return json.dumps(obj, ensure_ascii=False)


def casos(D: dict):
    mar, aer = D.get("maritimo", {}), D.get("aereo", {})
    salida, alertas = D.get("salida_dc"), D.get("alertas_eta")
    yield ("ESC-06 marítimo arribado fuera de Transito sale con su buque",
           "SINTM-200002" in texto(mar) and "SINT PACIFICO" in texto(mar))
    yield ("ESC-08 alerta de ETA > 5 días para SINTM-200004",
           alertas is not None and "SINTM-200004" in texto(alertas))
    yield ("ESC-08 la alerta trae ETA anterior, nueva y delta",
           alertas is not None and any("SINTM-200004" in texto(a) and {"eta_anterior", "eta_nueva", "delta_dias"} <= set(a)
                                       for a in (alertas if isinstance(alertas, list) else alertas.get("items", []))))
    yield ("ESC-11 embarque con doble MOT entra al radar aéreo", "SINTA-100008" in texto(aer))
    yield ("ESC-03 aéreo con salida FWD en vista Salida a DC", salida is not None and "SINTA-100004" in texto(salida))
    yield ("ESC-07 marítimo con salida FWD en vista Salida a DC", salida is not None and "SINTM-200003" in texto(salida))
    yield ("ESC-04 aéreo entregado (3PL) en vista Salida a DC", salida is not None and "SINTA-100005" in texto(salida))
    yield ("ESC-12 433 con DN en M4 marcado ingresado en SAP",
           salida is not None and "4339900602" in texto(salida) and "ingresado" in texto(salida).lower())
    yield ("ESC-12 433 sin DN marcado pendiente SAP",
           salida is not None and "4339900601" in texto(salida) and "pendiente" in texto(salida).lower())
    yield ("Sello de actualización para la corrida horaria", bool(D.get("actualizado") or D.get("generado")))
    # lo que ya funciona no se puede romper
    yield ("NO ROMPER ESC-02A aéreo arribado fuera del bi", "SINTA-100002" in texto(aer))
    yield ("NO ROMPER ESC-05 buque cercano SINT AURORA", "SINT AURORA" in texto(mar))
    yield ("NO ROMPER ESC-10 422 con dos 433", "4339900401" in texto(aer) and "4339900402" in texto(aer))


def main() -> int:
    D = datos_radar(CARPETA)
    resultados = list(casos(D))
    for nombre, ok in resultados:
        print(f"{'PASA ' if ok else 'FALLA'}  {nombre}")
    fallas = sum(not ok for _, ok in resultados)
    print(f"\n{len(resultados) - fallas}/{len(resultados)} pasan · carpeta {CARPETA}")
    return 1 if fallas else 0


if __name__ == "__main__":
    sys.exit(main())
