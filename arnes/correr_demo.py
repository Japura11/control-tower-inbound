# -*- coding: utf-8 -*-
"""
correr_v2.py — Ejecuta EN PROCESO (sin subprocess) los módulos consolidados de v2/
sobre los mismos datos sintéticos y fakes que correr_base.py durante 2 días consecutivos
(Día 1 = ayer, Día 2 = hoy) y guarda TODAS las salidas en salida_v2/.

Orden por cada día (único cambio frente a correr_base.py: pedidos corre ANTES que correo;
además correo pasa armar() en memoria al radar eliminando R1):
  1. in_transit (ingesta.main_in_transit)
  2. eventos    (eventos.main_eventos)
  3. pedidos    (ingesta.main_pedidos)
  4. correo     (salidas.main_correo -> devuelve armado en memoria)
  5. backfill   (ingesta.main_backfill)
  6. radar      (salidas.main_radar(armado=armado))

Uso: python arnes/correr_v2.py [--hoy 2026-11-25] [--semilla 20261125]
"""

from __future__ import annotations

import argparse
import io
import logging
import os
import shutil
import sys
import tempfile
import traceback
from datetime import datetime, timedelta
from pathlib import Path

os.environ.setdefault("PYTHONIOENCODING", "utf-8")
for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

RAIZ_V2 = Path(__file__).resolve().parent.parent
DIR_MODULOS_V2 = RAIZ_V2 / "torre"
DIR_SINTETICOS = RAIZ_V2 / "sinteticos"
DIR_SANDBOX = RAIZ_V2 / "sandbox"
DIR_SALIDA_V2 = RAIZ_V2 / "salida_demo"

for p in (str(DIR_SINTETICOS), str(DIR_MODULOS_V2)):
    if p not in sys.path:
        sys.path.insert(0, p)

from fakes import aplicar_reloj_a_modulos, instalar_fakes  # noqa: E402
from generar import (  # noqa: E402
    DIR_INBOUND,
    DIR_M4,
    DIR_TRACKING,
    FECHA_HOY_DEFAULT,
    SEMILLA_DEFAULT,
    construir_escenarios_y_datos,
)


class _TeeStream:
    """Duplica stdout/stderr hacia un buffer en memoria además de la consola."""

    def __init__(self, original, buffer: io.StringIO):
        self._orig = original
        self._buf = buffer

    def write(self, s: str) -> int:
        self._buf.write(s)
        return self._orig.write(s)

    def flush(self) -> None:
        self._buf.flush()
        self._orig.flush()


def copiar_si_existe(origen: Path, destino: Path) -> None:
    if origen.exists():
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(origen, destino)


def limpiar_salida_v2() -> None:
    DIR_SALIDA_V2.mkdir(parents=True, exist_ok=True)
    for item in DIR_SALIDA_V2.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        elif item.suffix.lower() in {".xlsx", ".parquet", ".html", ".json", ".log"}:
            item.unlink()


def correr_cadena(fecha_hoy_str: str = FECHA_HOY_DEFAULT, semilla: int = SEMILLA_DEFAULT) -> int:
    limpiar_salida_v2()
    estado = construir_escenarios_y_datos(fecha_hoy_str=fecha_hoy_str, semilla=semilla)

    # Asegurar temporales dentro de sandbox/_tmp
    dir_tmp = DIR_SANDBOX / "_tmp"
    dir_tmp.mkdir(parents=True, exist_ok=True)
    os.environ["TMP"] = os.environ["TEMP"] = str(dir_tmp)
    tempfile.tempdir = str(dir_tmp)

    # Inyectar todos los fakes ANTES de importar cualquier módulo de v2
    instalar_fakes(estado)

    import eventos
    import ingesta
    import salidas
    import torre
    import torre_config
    import vista

    # Guardar caché y mapa del radar en sandbox/Tracking/ igual que correr_base.py
    torre_config.RUTA_CATALOGO = DIR_TRACKING / "catalogo_buques.json"
    torre_config.RUTA_MAPA = DIR_TRACKING / "mapa_mundial_50m.json"
    salidas.PAUSA_S = (0, 0)
    torre_config.DRY_RUN_CORREO = True

    modulos_con_reloj = [
        torre_config,
        ingesta,
        eventos,
        vista,
        salidas,
        torre,
    ]
    aplicar_reloj_a_modulos(modulos_con_reloj, estado.reloj)

    hoy = datetime.strptime(fecha_hoy_str, "%Y-%m-%d")
    ayer = hoy - timedelta(days=1)
    dias_corrida = [
        (1, "dia1", ayer.replace(hour=8, minute=30, second=0)),
        (2, "dia2", hoy.replace(hour=8, minute=30, second=0)),
    ]

    fallos_ejecucion: list[dict] = []
    log_global = io.StringIO()

    # Conectar logging estándar de Python al stream capturado
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)
    handler_consola = logging.StreamHandler(sys.stdout)
    handler_consola.setFormatter(logging.Formatter("%(asctime)s %(name)-16s %(levelname)-7s %(message)s", "%H:%M:%S"))

    for num_dia, etiqueta_dia, momento in dias_corrida:
        dir_dia = DIR_SALIDA_V2 / etiqueta_dia
        dir_dia.mkdir(parents=True, exist_ok=True)
        estado.activar_dia(num_dia, momento, dir_dia)

        # Ajustar mtime previo de OM.xlsx 5 segundos antes del corte del día para el assert de refrescar_reporte
        ts_prev = momento.timestamp() - 5.0
        os.utime(DIR_INBOUND / "OM.xlsx", (ts_prev, ts_prev))

        buf_dia = io.StringIO()
        stdout_real, stderr_real = sys.stdout, sys.stderr
        sys.stdout = _TeeStream(stdout_real, buf_dia)
        sys.stderr = _TeeStream(stderr_real, buf_dia)
        handler_consola.stream = sys.stdout
        root_logger.addHandler(handler_consola)

        armado_dia = None
        pasos_dia = [
            "in_transit",
            "eventos",
            "pedidos",
            "correo",
            "backfill",
            "radar",
        ]

        try:
            print(f"\n{'═' * 72}")
            print(f"CORRIDA V2 · DÍA {num_dia} ({etiqueta_dia.upper()}) · RELOJ = {momento:%Y-%m-%d %H:%M:%S}")
            print(f"{'═' * 72}")
            for nombre_paso in pasos_dia:
                print(f"\n──── [{etiqueta_dia}] Paso: {nombre_paso} ────")
                try:
                    if nombre_paso == "correo":
                        armado_dia = torre.ejecutar_paso("correo")
                    elif nombre_paso == "radar":
                        torre.ejecutar_paso("radar", armado=armado_dia)
                    else:
                        torre.ejecutar_paso(nombre_paso)

                    # Alinear mtime de los archivos de salida con el reloj virtual del día
                    for arch in (DIR_INBOUND / "in_transit_peru.xlsx", DIR_INBOUND / "eventos_transito.xlsx"):
                        if arch.exists():
                            ts_arch = momento.timestamp()
                            os.utime(arch, (ts_arch, ts_arch))
                except SystemExit as exc:
                    if exc.code not in (None, 0):
                        tb = traceback.format_exc()
                        fallos_ejecucion.append({
                            "dia": num_dia,
                            "paso": nombre_paso,
                            "error": f"SystemExit({exc.code})",
                            "traceback": tb,
                        })
                        print(f"FALLO SystemExit en {nombre_paso}: {exc}")
                except Exception as exc:
                    tb = traceback.format_exc()
                    fallos_ejecucion.append({
                        "dia": num_dia,
                        "paso": nombre_paso,
                        "error": f"{type(exc).__name__}: {exc}",
                        "traceback": tb,
                    })
                    print(f"FALLO {type(exc).__name__} en {nombre_paso}: {exc}\n{tb}")
        finally:
            root_logger.removeHandler(handler_consola)
            sys.stdout, sys.stderr = stdout_real, stderr_real

        texto_log_dia = buf_dia.getvalue()
        (dir_dia / "corrida.log").write_text(texto_log_dia, encoding="utf-8")
        log_global.write(texto_log_dia)

        # Copiar todos los artefactos del día a salida_v2/diaN/
        for nombre_arch in (
            "OM.xlsx",
            "in_transit_peru.xlsx",
            "eventos_transito.xlsx",
            "historial_transito.parquet",
            "historial_pedidos.parquet",
        ):
            copiar_si_existe(DIR_INBOUND / nombre_arch, dir_dia / nombre_arch)
        copiar_si_existe(DIR_TRACKING / "radar_embarques.html", dir_dia / "radar_embarques.html")
        copiar_si_existe(DIR_TRACKING / "catalogo_buques.json", dir_dia / "catalogo_buques.json")

    # Guardar las salidas finales (Día 2 = "hoy") en la raíz de salida_v2/
    dir_final = DIR_SALIDA_V2 / "dia2"
    for archivo in dir_final.iterdir():
        if archivo.is_file() and archivo.name != "corrida.log":
            shutil.copy2(archivo, DIR_SALIDA_V2 / archivo.name)
    for m4_arch in ("m4_latam.parquet", "m4_latam_sin_gi.parquet"):
        copiar_si_existe(DIR_M4 / m4_arch, DIR_SALIDA_V2 / m4_arch)

    (DIR_SALIDA_V2 / "corrida_completa.log").write_text(log_global.getvalue(), encoding="utf-8")

    print(f"\n{'═' * 72}")
    print(f"CORRIDA V2 COMPLETADA -> {DIR_SALIDA_V2}")
    print(f"Fallos no controlados: {len(fallos_ejecucion)}")
    escaneado = [f for f in estado.lecturas_pdf if "escaneado" in f]
    print(f"Lecturas de PDF escaneado en la corrida: {len(escaneado)} {escaneado}")
    print(f"{'═' * 72}")
    return 1 if fallos_ejecucion else 0


def main() -> None:
    parser = argparse.ArgumentParser(description="Corre los módulos v2 en proceso (Día 1 y Día 2) con fakes.")
    parser.add_argument("--hoy", default=FECHA_HOY_DEFAULT, help="Fecha 'hoy' YYYY-MM-DD (default: 2026-11-25)")
    parser.add_argument("--semilla", type=int, default=SEMILLA_DEFAULT, help="Semilla determinista")
    args = parser.parse_args()
    sys.exit(correr_cadena(fecha_hoy_str=args.hoy, semilla=args.semilla))


if __name__ == "__main__":
    main()
