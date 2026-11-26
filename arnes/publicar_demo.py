"""Copia las salidas del demo a demo/ (servidas con GitHub Pages) y las limpia: sin rutas locales y sin nada fuera
de lo sintético. Falla si encuentra algo sospechoso.
Uso: python arnes/correr_demo.py && python arnes/demo_varios.py && python arnes/publicar_demo.py"""
import json
import re
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
DEMO = RAIZ / "demo"
REPO = "https://github.com/Japura11/control-tower-inbound#readme"
SOSPECHOSO = re.compile(r"file:///|[A-Za-z]:[\\/]Users|lenovo|rcruz|onedrive|kuehne|vesselfinder|flexport|\bdhl\b|\bkudu\b",
                        re.I)


def limpiar_radar(html: str) -> str:
    m = re.search(r"const D = (\{.*?\});\s*\n", html, re.S)
    D = json.loads(m.group(1))
    D["link_detalle"], D["nombre_detalle"] = REPO, "repositorio · el Excel de detalle no viaja en el demo"
    return html[:m.start(1)] + json.dumps(D, ensure_ascii=False) + html[m.end(1):]


def main() -> None:
    DEMO.mkdir(exist_ok=True)
    salidas = {
        "radar.html": limpiar_radar((RAIZ / "salida_demo" / "demo_varios.html").read_text(encoding="utf-8")),
        "correo.html": (RAIZ / "salida_demo" / "dia2" / "correo_arribados.html").read_text(encoding="utf-8"),
    }
    for nombre, html in salidas.items():
        malos = sorted(set(SOSPECHOSO.findall(html)))
        assert not malos, f"{nombre}: quedan rastros {malos}"
        (DEMO / nombre).write_text(html, encoding="utf-8", newline="\n")
    (DEMO / "index.html").write_text("""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Control Tower Inbound · demo</title>
<style>body{font:16px/1.5 'Segoe UI',Arial,sans-serif;margin:0;background:#0F1B3D;color:#fff;display:grid;place-items:center;min-height:100vh}
main{max-width:640px;padding:32px}h1{margin:0 0 4px}p{color:#C8D6EE}a{display:block;margin:14px 0;padding:16px 18px;border-radius:10px;
background:#fff;color:#1A1726;text-decoration:none;border-left:6px solid #5B3E96}a:hover{border-left-color:#E4572E}
a b{display:block;font-size:18px}small{color:#4A4658}</style></head><body><main>
<h1>Control Tower Inbound</h1><p>Demo con datos 100&nbsp;% sintéticos. Nada aquí es de una empresa real.</p>
<a href="radar.html"><b>Radar de embarques →</b><small>Marítimo (rama y carta), Aéreo, Salida a DC con estado SAP y Alertas de ETA. Busca por buque, BL, 422 o 433.</small></a>
<a href="correo.html"><b>Correo de arribados →</b><small>Acción primero: fuera de SLA y bultos malos, luego Arribado y En ruta al DC.</small></a>
<a href="https://github.com/Japura11/control-tower-inbound"><b>Código y documentación →</b><small>Marco de trabajo, problemas y jugadas, pruebas.</small></a>
</main></body></html>
""", encoding="utf-8", newline="\n")
    print("demo/ lista:", sorted(p.name for p in DEMO.iterdir()))


if __name__ == "__main__":
    main()
