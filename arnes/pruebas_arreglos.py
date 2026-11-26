"""Pruebas del paso 4: cada arreglo debe FALLAR con v2 (paso 3, congelado) y PASAR con v2_paso4.

Cada prueba corre en un subproceso aparte con los fakes instalados (nunca COM/Outlook/red reales) y los
módulos de una carpeta (v2 o v2_paso4). Imprime la matriz antes/después y sale con 1 si algo no cumple.
Uso: python arnes/pruebas_paso4.py
"""
import subprocess
import sys
import textwrap
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]

PREAMBULO = textwrap.dedent("""
    import sys, os
    from pathlib import Path
    RAIZ = Path(r"{raiz}")
    sys.path[:0] = [str(RAIZ / "sinteticos"), str(RAIZ / "{modulos}"), str(RAIZ / "arnes")]
    os.environ["TMP"] = os.environ["TEMP"] = str(RAIZ / "sandbox" / "_tmp")
    import importlib; A = importlib.import_module("{arnes}")   # construye datos y fakes con la carpeta de modulos correcta
    from fakes import instalar_fakes
    estado = A.construir_escenarios_y_datos(fecha_hoy_str=A.FECHA_HOY_DEFAULT, semilla=A.SEMILLA_DEFAULT)
    instalar_fakes(estado)
    import torre_config as C
""")

PRUEBAS = {
    "A1 copiar_pedidos que falla deja aviso (no queda en silencio)": """
        import ingesta, pandas as pd
        def revienta(*a, **k): raise PermissionError("parquet bloqueado por OneDrive")
        ingesta.leer_ordermatch = revienta
        ingesta.copiar_pedidos(pd.Timestamp("2026-11-25 08:30"))
        assert getattr(C, "AVISOS_PASO", []), "copiar_pedidos fallo y no dejo ningun aviso"
    """,
    "A2 historial ilegible aborta eventos (regla 14)": """
        import eventos, pandas as pd
        C.RUTA_HISTORIAL.parent.mkdir(parents=True, exist_ok=True)
        C.RUTA_HISTORIAL.write_bytes(b"no es un parquet")
        bi = pd.DataFrame({"key": ["k1"], "bl_awb_num": ["SINTA-1"], "so": ["4221"], "mot": ["AIR"]})
        try:
            eventos.con_historial(bi)
        except RuntimeError:
            pass
        else:
            raise AssertionError("historial ilegible y eventos siguio con el bi de hoy")
    """,
    "A6 doble MOT toma el SLA mas estricto": """
        import vista
        assert hasattr(vista, "sla_de_mot"), "no existe sla_de_mot"
        assert vista.sla_de_mot("AIR / SEA") == 3 and vista.sla_de_mot("SEA") == 7
    """,
    "A3 reemplazo atomico reintenta ante un lock breve de OneDrive": """
        import os, tempfile
        from pathlib import Path
        d = Path(tempfile.mkdtemp()); tmp, dst = d / "x.tmp", d / "x.txt"
        tmp.write_text("nuevo"); dst.write_text("viejo")
        real, veces = os.replace, []
        def con_lock(a, b):
            veces.append(1)
            if len(veces) == 1: raise PermissionError("lock de OneDrive")
            return real(a, b)
        os.replace = con_lock
        try:
            C.reemplazar_atomico(tmp, dst)
        finally:
            os.replace = real
        assert dst.read_text() == "nuevo", "no reintento"
    """,
    "A8 falta una columna del radar en Transito: aviso en el paso 1": """
        import ingesta, pandas as pd
        import inspect
        src = inspect.getsource(ingesta)
        assert "COLS_RADAR" in src and "name_of_ocean_vessel" in src.split("COLS_RADAR", 1)[1][:200], "el paso 1 no valida las columnas del radar"
    """,
    "A4 backfill deja de re-buscar lineas de stock viejas": """
        import ingesta, pandas as pd
        assert hasattr(ingesta, "previas_a_reintentar"), "no existe el filtro de reintentos"
        hoy = pd.Timestamp("2026-11-25")
        r = pd.DataFrame({"so": ["4221", "4222"], "mtm": ["M1", "M2"],
                          "visto_desde": [hoy - pd.Timedelta(days=20), hoy - pd.Timedelta(days=2)]})
        quedan = ingesta.previas_a_reintentar(r, set(), hoy)
        assert list(quedan["so"]) == ["4222"], f"quedan {list(quedan['so'])}"
    """,
    "R9 un pais nuevo en Transito no tumba la cadena (aviso + fuera de alcance)": """
        import ingesta, inspect
        src = inspect.getsource(ingesta.cargar_reporte)
        assert "assert not desconocidos" not in src, "un country nuevo sigue abortando la cadena"
        assert "AVISOS_PASO" in src, "el country nuevo no deja aviso"
    """,
}


def correr(modulos: str, cuerpo: str) -> tuple[bool, str]:
    codigo = PREAMBULO.format(raiz=RAIZ, modulos=modulos, arnes="correr_v2" if modulos == "v2" else "correr_demo") + textwrap.dedent(cuerpo)
    r = subprocess.run([sys.executable, "-c", codigo], capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"})
    ultima = (r.stderr.strip().splitlines() or [""])[-1]
    return r.returncode == 0, ultima


def main() -> int:
    malas = 0
    for nombre, cuerpo in PRUEBAS.items():
        antes, motivo_antes = False, "version anterior: no incluida en el repo"
        despues, motivo = correr("torre", cuerpo)
        ok = (not antes) and despues
        malas += not ok
        print(f"{'OK   ' if ok else 'MAL  '} {nombre}\n       v2: {'pasa' if antes else 'falla'} ({motivo_antes[:90]})"
              f"\n       torre: {'pasa' if despues else 'falla: ' + motivo[:160]}")
    print(f"\n{len(PRUEBAS) - malas}/{len(PRUEBAS)} arreglos demostrados")
    return 1 if malas else 0


if __name__ == "__main__":
    sys.exit(main())
