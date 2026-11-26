# -*- coding: utf-8 -*-
"""
torre.py — Orquestador y chequeo de la cadena Control Tower Perú (v2 · consolidación de orquestador_transito.py
y chequeo_cadena.py).

Modos de uso:
  python torre.py diario          corre en subprocesos aislados: in_transit -> eventos -> pedidos -> correo -> radar
                                  (pedidos corre ANTES que correo; correo y radar comparten armar() en memoria · R1)
  python torre.py backfill        rescata líneas y 433 desde correos enviados «Inbound & Order Match»
  python torre.py radar           recalcula únicamente radar_embarques.html sobre la base ya escrita
  python torre.py chequeo         audita de punta a punta scripts, historiales, seguimiento, eventos, Outlook y correo
  python torre.py paso <nombre>   ejecuta un paso individual en este proceso (aislamiento COM por subproceso)

El arnés de pruebas puede invocar `correr_paso_en_proceso()` o las funciones `main_*` directamente en memoria.
"""

import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

import torre_config as C

# ══ PARA CORRER DIRECTO (boton Run del editor o «python torre.py» sin nada mas) ══
MODO     = "diario"      # "diario" | "radar_horario" | "semanal" | "chequeo" | "backfill" | "radar"
SIN_LAKE = False         # True: la corrida diaria no consulta Lake (mas rapida; sin bultos/kg/m3 de Lake)

SCRIPT_TORRE = Path(__file__).resolve()
SCRIPTS      = SCRIPT_TORRE.parent

# Orden diario v2: pedidos ANTES que correo (pedidos_locales no depende del correo);
# correo_radar ejecuta correo -> radar en un mismo subproceso pasando armar() en memoria (R1).
PASOS_DIARIO_SUBPROCESO = [
    ("in_transit", "in_transit"),
    ("eventos", "eventos"),
    ("pedidos", "pedidos"),
    ("correo+radar", "correo_radar"),
]
PASOS_NOMBRES_DIARIO = ("in_transit", "eventos", "pedidos", "correo", "radar")

# ══════════════════════════ CONFIG · CHEQUEO DE CADENA ═══════════════════════
REVISAR_OUTLOOK = True
MARCAS_V2 = {
    "torre_config.py": [
        ("USAR_LAKE", "declara interruptor USAR_LAKE"),
        ("def sold_to_por_pedido(", "muestra el sold-to junto al cliente"),
        ("def leer_embarques(", "lee embarques convertido"),
    ],
    "ingesta.py": [
        ("def actualizar_historial(", "guarda el historial"),
        ("def copiar_pedidos(", "guarda la copia diaria de los pedidos del OrderMatch"),
        ("def columna_sold_to(", "guarda el sold-to de cada pedido"),
        ("def rescatar(", "busca en los Inbound el HAWB de los avisos de esta semana"),
        ("def pedidos_del_inbound(", "trae el 433 del Order Match del Inbound"),
        ("def salidas_sin_linea(", "trae tambien lo que tiene salida a DC"),
        ("EXT_ADJUNTO", "lee tambien el adjunto Order Match del correo"),
        ("previas_con_sold_to", "trae el sold-to del Order Match"),
        ("RESULTADO:", "dice al final que hizo (RESULTADO)"),
    ],
    "eventos.py": [
        ("def resolver(", "correos FWD"),
        ("def con_historial(", "cruza avisos con el historial"),
        ("desde_salidos", "busca 30 dias atras el aviso de lo que salio del bi"),
        ("se registra igual, con la fecha del correo", "registra avisos sin PDF legible"),
    ],
    "vista.py": [
        ("def lineas_que_salieron(", "muestra los aereos que salieron del in-transit"),
        ("def pedidos_para_salidos(", "les da el 433 aunque el OrderMatch de hoy ya no lo traiga"),
        ("maritimo: solo con salida a DC", "muestra los maritimos que salieron con ruta a DC"),
        ("def aereos_en_terminal(", "reutiliza armar() en memoria para el radar"),
    ],
    "salidas.py": [
        ("def cuerpo_html(", "renderiza el correo de arribados"),
        ("def escribir_html(", "escribe radar_embarques.html"),
        ("def posicion_proveedor_b(", "consulta AIS de buques"),
    ],
}


# ══════════════════════════ CHEQUEO DE CADENA (chequeo_cadena.py) ════════════
def _ok(texto: str) -> None:
    print(f"   OK     {texto}")


def _nota(texto: str) -> None:
    print(f"   NOTA   {texto}")


def _falla(fallas: list, paso: int, texto: str, arreglo: str) -> None:
    print(f"   FALLA  {texto}\n          arreglo: {arreglo}")
    fallas.append((paso, texto, arreglo))


def _leer_script(nombre: str) -> str | None:
    ruta = SCRIPTS / nombre
    if ruta.exists():
        return ruta.read_text(encoding="utf-8", errors="replace")
    return None


def _paso_scripts(fallas: list) -> None:
    print(f"\n1) Módulos v2 en {SCRIPTS}")
    for nombre, marcas in MARCAS_V2.items():
        texto = _leer_script(nombre)
        if texto is None:
            _falla(fallas, 1, f"{nombre}: no esta en la carpeta", f"copiar {nombre} a {SCRIPTS}")
            continue
        faltan = [que for marca, que in marcas if marca not in texto]
        if faltan:
            _falla(fallas, 1, f"{nombre}: version incompleta (no {'; no '.join(faltan)})",
                   f"actualizar {nombre} a la version actual")
        else:
            _ok(f"{nombre}: version actual")
    cfg = _leer_script("torre_config.py") or ""
    for var_cfg, etiqueta, en_seco in (
        ("DRY_RUN_EVENTOS", "eventos.py", "no registra avisos ni escribe seguimiento"),
        ("DRY_RUN_CORREO", "salidas.py (correo)", "abre el borrador sin enviar"),
    ):
        m = re.search(rf"^{var_cfg}\s*=\s*(True|False)\s*(?:#.*)?\r?$", cfg, re.M)
        seco = m is not None and m.group(1) == "True"
        if seco and var_cfg == "DRY_RUN_EVENTOS":
            _falla(fallas, 1, f"{etiqueta} en DRY_RUN = True: {en_seco}",
                   "poner DRY_RUN_EVENTOS = False en torre_config.py y correr torre.py diario")
        else:
            _ok(f"{etiqueta}: {var_cfg} = {m.group(1) if m else '?'}" + (f" ({en_seco})" if seco else ""))
    print(f"   torre.py diario: {' -> '.join(PASOS_NOMBRES_DIARIO)}")


def _paso_avisos_semana(h: pd.DataFrame, fallas: list) -> None:
    """Lo que haria el backfill, sin escribir: avisos de esta semana con HAWB fuera del bi y del historial."""
    if not REVISAR_OUTLOOK:
        return
    try:
        import ingesta as I
        nuevas, informe, r = I.rescatar(h)
    except Exception as e:
        _falla(fallas, 2, f"no se pudo revisar los avisos de esta semana ({type(e).__name__}: {e})",
               "correr python torre.py backfill y pegar su salida")
        return
    if len(nuevas):
        _falla(fallas, 2,
               f"{nuevas['bl_awb_num'].nunique()} aereos con aviso de esta semana estan en un Inbound pero no en el "
               f"historial: {', '.join(nuevas['bl_awb_num'].unique()[:8])}",
               "correr aparte: python torre.py backfill, y despues python torre.py diario")
    elif len(r.get("pedidos", [])):
        _falla(fallas, 2,
               f"{len(r['pedidos'])} pedidos (433) de lineas traidas de los Inbound estan en su Order Match pero no en "
               "historial_pedidos.parquet: el correo no las muestra",
               "correr aparte: python torre.py backfill, y despues python torre.py diario")
    elif r["avisos"] or r.get("salidas"):
        _nota(f"{r['avisos']} avisos y {r.get('salidas', 0)} salidas a DC de esta semana sin su HAWB/BL en el bi, el "
              f"historial ni los Inbound (no hay de donde sacar su SO y MTM): {', '.join(informe['hawb'].head(8))}")
    elif r["cubiertos"]:
        _ok(f"los avisos de esta semana tienen su HAWB en el bi o el historial ({r['cubiertos']} HAWB)")
    else:
        import ingesta as I
        _nota(f"no hay avisos de llegada con HAWB en los ultimos {I.DIAS_AVISOS} dias")


def _paso_historial(ruta: Path, bi: pd.DataFrame, corte_bi, fallas: list) -> pd.DataFrame:
    print(f"\n2) {ruta.name}")
    if not ruta.exists():
        _falla(fallas, 2, "no existe", "correr python torre.py diario (ingesta lo crea)")
        return pd.DataFrame()
    h = pd.read_parquet(ruta)
    fuente = h["fuente"].astype(str).str.replace(r"inbound semanal .*", "inbound semanal", regex=True).value_counts()
    print(f"   {len(h):,} filas · {h['key'].nunique():,} lineas · {h['visto_desde'].min():%d/%m %H:%M} -> "
          f"{h['visto_hasta'].max():%d/%m %H:%M} · fuentes {fuente.to_dict()}")
    pedidos = C.RUTA_HISTORIAL_PEDIDOS
    if pedidos.exists():
        hp = pd.read_parquet(pedidos)
        fuentes = hp["fuente"].astype(str).str.replace(r"inbound semanal .*", "inbound semanal", regex=True)
        _ok(f"{pedidos.name}: {len(hp):,} 422 x MTM -> 433 · al {hp['visto_hasta'].max():%d/%m %H:%M} · "
            f"fuentes {fuentes.value_counts().to_dict()}")
    else:
        _falla(fallas, 2, f"no existe {pedidos.name}: lo que sale del in-transit no tiene de donde sacar su 433",
               "correr python torre.py diario y python torre.py backfill")
    _paso_avisos_semana(h, fallas)
    if pd.notna(corte_bi) and abs(h["visto_hasta"].max() - corte_bi) > pd.Timedelta(hours=2):
        _falla(fallas, 2,
               f"no tiene el ultimo bi (historial al {h['visto_hasta'].max():%d/%m %H:%M}, bi del {corte_bi:%d/%m %H:%M})",
               "correr python torre.py diario para actualizar el historial")
    ultima = h.sort_values(["visto_desde", "visto_hasta"]).groupby("key").tail(1)
    fuera = ultima[~ultima["key"].astype(str).isin(set(bi["key"].astype(str)))
                   & pd.to_datetime(ultima["visto_hasta"]).ge(datetime.now() - timedelta(days=C.DIAS_VENTANA_SALIDOS))]
    salidos = fuera[~fuera["mot"].astype(str).str.upper().str.strip().isin(C.MOT_MARITIMO)].dropna(subset=["bl_awb_num"])
    salidos = salidos.assign(hawb_key=salidos["bl_awb_num"].map(C.normalizar_id))
    if salidos.empty:
        _nota(f"ningun aereo salio del bi en los ultimos {C.DIAS_VENTANA_SALIDOS} dias segun el historial (empieza el "
              f"{h['visto_desde'].min():%d/%m})")
    else:
        _ok(f"{salidos['hawb_key'].nunique()} aereos (HAWB) salieron del bi en los ultimos {C.DIAS_VENTANA_SALIDOS} dias")
    return salidos


def _paso_seguimiento(ruta: Path, fallas: list):
    print(f"\n3) seguimiento en {ruta.name}")
    seg = pd.read_excel(ruta, "seguimiento")
    corte = pd.Timestamp(seg["corte_bi"].iloc[0])
    if "atd_correo" not in seg.columns:
        _falla(fallas, 3, "lo escribio el eventos anterior (sin columnas de correos FWD)",
               "correr eventos actual con DRY_RUN_EVENTOS = False")
    else:
        _ok(f"lo escribio el eventos actual · hitos {seg['hito'].value_counts().to_dict()}")
    if datetime.now() - corte > timedelta(hours=C.MAX_HORAS):
        _falla(fallas, 3, f"corte del {corte:%d/%m %H:%M}: tiene mas de {C.MAX_HORAS} h y el correo no corre",
               "correr la cadena completa")
    return corte


def _paso_eventos(ruta: Path, salidos: pd.DataFrame, fallas: list):
    print(f"\n4) {ruta.name}")
    with pd.ExcelFile(ruta) as libro:
        hojas = libro.sheet_names
        arribos = pd.read_excel(libro, "arribos", dtype={"hawb_key": str})
        salidas = pd.read_excel(libro, "salidas_dc", dtype={"hawb_key": str})
    respaldos = sorted(ruta.parent.glob(f"{ruta.stem}_respaldo_*.xlsx"))
    if "notificaciones" not in hojas:
        _falla(fallas, 4, "el eventos actual nunca corrio en real (no hay hoja notificaciones ni respaldo)",
               "DRY_RUN_EVENTOS = False en torre_config.py y correr python torre.py diario")
    else:
        _ok(f"el eventos actual ya corrio en real · respaldo {respaldos[-1].name if respaldos else 'no encontrado'}")
    ultimo = lambda df, col: (df.dropna(subset=["hawb_key"]).sort_values("fecha_mail")
                                .drop_duplicates("hawb_key", keep="last").set_index("hawb_key")[col])
    aviso, salida = ultimo(arribos, "fecha_mail"), ultimo(salidas, "fecha_programada")
    if len(salidos):
        con = salidos["hawb_key"].isin(aviso.index)
        texto = (f"{salidos.loc[con, 'hawb_key'].nunique()} de {salidos['hawb_key'].nunique()} aereos que salieron del bi "
                 "tienen aviso registrado")
        if con.any():
            _ok(texto)
        else:
            _falla(fallas, 4, texto, "ver el paso 5: si el aviso esta en Outlook, eventos no lo leyo")
    return aviso, salida


def _paso_outlook(sin_aviso: list, fallas: list) -> dict:
    print(f"\n5) Outlook: avisos de los {len(sin_aviso)} aereos sin aviso registrado")
    if not sin_aviso or not REVISAR_OUTLOOK:
        print("   nada que buscar")
        return {}
    import win32com.client as win32
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    filtro = f"[ReceivedTime] >= '{datetime.now() - timedelta(days=C.DIAS_VENTANA_SALIDOS):%m/%d/%Y %I:%M %p}'"
    hallados = {}
    for carpeta in C.carpetas_outlook(ns):
        for item in carpeta.Items.Restrict(filtro):
            try:
                asunto = getattr(item, "Subject", "") or ""
                if getattr(item, "Class", 0) != 43 or not C.ASUNTO_ARRIBO.search(asunto):
                    continue
                plano = C.normalizar_id(f"{asunto} {getattr(item, 'Body', '') or ''}")
                for k in sin_aviso:
                    if k in plano:
                        recibido = item.ReceivedTime.replace(tzinfo=None)
                        en_asunto = k in C.ids_en_asunto(asunto)
                        motivo = ("es respuesta o reenvio (RE/RV/FW): eventos no toma respuestas como aviso"
                                  if C.es_respuesta(asunto) else
                                  "el HAWB no esta en el asunto: eventos lo busca solo ahi" if not en_asunto else
                                  "el aviso esta bien: eventos lo toma en su proxima corrida (antes saltaba los "
                                  "avisos con PDF escaneado)")
                        hallados.setdefault(k, []).append((recibido, asunto, en_asunto, motivo))
            except Exception:
                continue
    if hallados:
        for k, avisos in hallados.items():
            recibido, asunto, _, motivo = max(avisos)
            print(f"   {k}: aviso del {recibido:%d/%m %H:%M} · {motivo} · {asunto[:70]}")
        _falla(fallas, 5, f"{len(hallados)} aereos tienen aviso en Outlook pero eventos no lo registro",
               "poner DRY_RUN_EVENTOS = False en torre_config.py y correr python torre.py diario" if C.DRY_RUN_EVENTOS else
               "el motivo de cada uno esta arriba; si dice 'el aviso esta bien', correr python torre.py diario y, "
               "si sigue, buscar ese HAWB en el log de eventos")
    else:
        _ok("ninguno tiene aviso en Outlook: no hay como saber que llegaron")
    return hallados


def _paso_correo(fallas: list) -> pd.DataFrame:
    print("\n6) Correo (armar sin enviar)")
    import vista as V
    tabla, _, _ = V.armar()
    fuera = tabla[tabla["fuera_in_transit"]] if len(tabla) else tabla
    if fuera.empty:
        _falla(fallas, 6, "el correo no trae nada fuera del in-transit",
               "ver la tabla por aereo: dice donde se cae cada uno")
    else:
        emb = fuera.drop_duplicates("embarque")
        _ok(f"el correo trae {len(emb)} embarques fuera del in-transit (por mot {emb['mot'].value_counts().to_dict()}): "
            f"{emb[['embarque', 'estado']].values.tolist()}")
    return tabla


def _veredicto(fallas: list) -> None:
    print("\nDonde falla")
    if fallas:
        paso, que, arreglo = fallas[0]
        print(f"   primer fallo, paso {paso}: {que}\n   arreglo: {arreglo}")
    else:
        print("   en ningun paso: los aereos que salieron del bi y tienen aviso estan en el correo")


def main_chequeo() -> None:
    fallas: list = []
    print(f"chequeo {datetime.now():%d/%m/%Y %H:%M}")
    _paso_scripts(fallas)
    try:
        import ingesta as I
        import vista as V
    except Exception as e:
        _falla(fallas, 1, f"los modulos v2 no se pueden importar ({type(e).__name__}: {e})",
               "revisar instalacion e imports en v2")
        return _veredicto(fallas)
    print(f"   datos en {C.CARPETA}")
    bi = pd.read_excel(C.RUTA_IN_TRANSIT, "bi", dtype={"key": str, "bl_awb_num": str})
    try:
        corte = pd.Timestamp(pd.read_excel(C.RUTA_IN_TRANSIT, "seguimiento", usecols=["corte_bi"], nrows=1).iloc[0, 0])
    except Exception:
        corte = None
    salidos = _paso_historial(C.RUTA_HISTORIAL, bi, corte, fallas)
    try:
        _paso_seguimiento(C.RUTA_IN_TRANSIT, fallas)
    except Exception as e:
        _falla(fallas, 3, f"no se pudo leer seguimiento ({type(e).__name__}: {e})", "correr python torre.py paso eventos")
    aviso, salida = _paso_eventos(C.RUTA_EVENTOS, salidos, fallas)
    sin_aviso = sorted(set(salidos["hawb_key"]) - set(aviso.index)) if len(salidos) else []
    hallados = _paso_outlook(sin_aviso, fallas)
    tabla = pd.DataFrame()
    try:
        tabla = _paso_correo(fallas)
    except Exception as e:
        _falla(fallas, 6, f"el correo no se pudo armar ({type(e).__name__}: {e})",
               "corregir lo que dice el mensaje y re-correr")

    if len(salidos):
        en_correo = (tabla.drop_duplicates("embarque").set_index("embarque")["estado"].rename(lambda x: C.normalizar_id(x))
                     if len(tabla) else pd.Series(dtype=str))
        try:
            del_om = I.pedidos_de_om(I.leer_ordermatch())
            del_om = V.pedidos_para_salidos(del_om)
            pedidos = I.lineas_con_pedido(salidos.drop(columns="hawb_key"), del_om)
            b2b = set(pedidos.loc[pedidos["vinculo"].eq("con pedido"), "bl_awb_num"].map(C.normalizar_id))
        except Exception as e:
            print(f"   (sin cruce con OrderMatch: {type(e).__name__}: {e})")
            b2b = set()
        desde_ruta = pd.Timestamp.now().normalize() - pd.offsets.BDay(C.DIAS_HABILES_ENTREGA)
        filas = []
        for k, g in salidos.groupby("hawb_key"):
            sal = pd.to_datetime(salida.get(k), errors="coerce")
            if k in en_correo.index:
                cae = f"sale en el correo ({en_correo[k]})"
            elif k not in aviso.index:
                h = hallados.get(k)
                cae = (f"sin aviso registrado; en Outlook el {max(h)[0]:%d/%m} ({max(h)[3]})" if h
                       else "sin aviso registrado ni en Outlook")
            elif pd.notna(sal) and sal < desde_ruta:
                cae = f"ya salio a DC el {sal:%d/%m}"
            elif k not in b2b:
                cae = "sin pedido B2B (solo stock o MTM distinto en OM)"
            else:
                cae = "revisar"
            filas.append({"hawb": g["bl_awb_num"].iloc[0], "salio_del_bi": f"{g['visto_hasta'].max():%d/%m}",
                          "aviso": "—" if k not in aviso.index else f"{pd.Timestamp(aviso[k]):%d/%m}",
                          "salida_fwd": "—" if pd.isna(sal) else f"{sal:%d/%m}", "se_cae_en": cae})
        print("\nAereos que salieron del bi: donde se cae cada uno")
        print(pd.DataFrame(filas).sort_values(["se_cae_en", "hawb"]).to_string(index=False))
    _veredicto(fallas)


# ══════════════════════════ EJECUCIÓN DE PASOS (PROCESO Y SUBPROCESO) ════════
def ejecutar_paso(nombre: str, armado=None):
    """Ejecuta un paso individual dentro del proceso actual."""
    if nombre == "in_transit":
        import ingesta
        ingesta.main_in_transit()
        return None
    if nombre == "eventos":
        import eventos
        eventos.main_eventos()
        return None
    if nombre == "pedidos":
        import ingesta
        ingesta.main_pedidos()
        return None
    if nombre == "correo":
        import salidas
        return salidas.main_correo(armado=armado)
    if nombre == "radar":
        import salidas
        salidas.main_radar(armado=armado)
        return None
    if nombre == "radar_horario":
        main_radar_horario()
        return None
    if nombre == "correo_radar":
        import salidas
        salidas.main_salidas()
        return None
    if nombre == "backfill":
        import ingesta
        ingesta.main_backfill()
        return None
    if nombre == "foto_semanal":
        foto_semanal()
        return None
    if nombre == "chequeo":
        main_chequeo()
        return None
    raise ValueError(f"Paso desconocido en torre.py: {nombre!r}")


def correr_subproceso(paso: str) -> str | None:
    """Corre `python torre.py paso <paso>` en un subproceso aislado (mantiene aislamiento COM Excel vs Outlook)."""
    assert SCRIPT_TORRE.exists(), f"No existe {SCRIPT_TORRE}"
    entorno = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
    ultima = ""
    with subprocess.Popen([sys.executable, str(SCRIPT_TORRE), "paso", paso], cwd=SCRIPTS, env=entorno,
                          stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                          encoding="utf-8", errors="replace") as proceso:
        for linea in proceso.stdout:
            print(linea, end="", flush=True)
            ultima = linea.strip() or ultima
    if proceso.returncode == C.SALIDA_CON_AVISO:                    # paso 4: siguio, pero se reporta
        return "AVISO " + ultima
    return None if proceso.returncode == 0 else ultima


def main_diario_subprocesos() -> None:
    resumen = []
    for etiqueta, paso_cmd in PASOS_DIARIO_SUBPROCESO:
        print(f"\n{'=' * 20} {etiqueta} {'=' * 20}", flush=True)
        t0 = time.perf_counter()
        error = correr_subproceso(paso_cmd)
        resumen.append((etiqueta, time.perf_counter() - t0, error))
        if error and not error.startswith("AVISO "):              # un aviso no detiene la cadena
            break
    print(f"\n{'=' * 50}")
    for etiqueta, segundos, error in resumen:
        print(f"   {etiqueta:<13} {segundos:>6.1f} s   {'OK' if error is None else 'FALLO -> ' + error}")
    print(f"   {'total':<13} {sum(s for _, s, _ in resumen):>6.1f} s")
    no_corrieron = [et for et, _ in PASOS_DIARIO_SUBPROCESO[len(resumen):]]
    if no_corrieron:
        print(f"   no corrieron: {no_corrieron}")
    if any(error for _, _, error in resumen):
        sys.exit(1)


def foto_semanal(hoy=None) -> Path:
    """Paso 5 (pedido 25-26/11): copia fija de los historiales y de eventos en fotos_semanales/AAAA-Wss. Si la foto de
    la semana ya existe NO se reescribe (es el respaldo de esa semana). Devuelve la carpeta."""
    import shutil
    hoy = pd_now() if hoy is None else hoy
    anio, semana, _ = hoy.isocalendar()
    destino = C.CARPETA_FOTOS / f"{anio}-W{semana:02d}"
    if destino.exists():
        print(f"   foto semanal {destino.name}: ya existe, no se reescribe", flush=True)
        return destino
    tmp = destino.with_name(destino.name + ".tmp")
    shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True)
    copiados = []
    for ruta in (C.RUTA_HISTORIAL, C.HIST_PEDIDOS_PATH, C.RUTA_EVENTOS):
        if ruta.exists():
            shutil.copy2(ruta, tmp / ruta.name)
            copiados.append(ruta.name)
    tmp.rename(destino)                                    # la foto aparece completa o no aparece
    print(f"   foto semanal {destino.name}: {copiados}", flush=True)
    return destino


def pd_now():
    from datetime import datetime
    return datetime.now()


def main_radar_horario() -> None:
    """Paso 5 (pedido 25/11): radar cada hora con lo que se mueve en el dia (avisos de llegada, salidas FWD,
    recepcion 3PL) SIN escribir la base. Copia in_transit_peru.xlsx y eventos_transito.xlsx a una carpeta temporal,
    corre eventos sobre las copias (lee solo lo nuevo de Outlook: su ventana se ancla en el ultimo correo
    registrado), arma el radar y borra las copias. El historial y los pedidos se leen, nunca se escriben: lo de la
    hora queda registrado de verdad en la corrida diaria. La base puede tener hasta C.MAX_HORAS_RADAR horas."""
    import shutil
    import tempfile
    import eventos
    import salidas
    tmp = Path(tempfile.mkdtemp(prefix="radar_horario_"))
    originales = {a: getattr(C, a) for a in ("RUTA_IN_TRANSIT", "RUTA_EVENTOS")}
    max_horas = C.MAX_HORAS
    try:
        for attr, ruta in originales.items():
            assert ruta.exists(), f"No existe {ruta} — correr primero: python torre.py diario"
            copia = tmp / ruta.name
            shutil.copy2(ruta, copia)
            setattr(C, attr, copia)
        C.MAX_HORAS = C.MAX_HORAS_RADAR
        C.HORAS_CACHE_AIS_VIGENTE = C.HORAS_CACHE_AIS_RADAR   # coordenadas frescas cada 2 h
        eventos.main_eventos()                      # escribe solo en las copias temporales
        salidas.main_radar()
    finally:
        for attr, ruta in originales.items():
            setattr(C, attr, ruta)
        C.MAX_HORAS = max_horas
        C.__dict__.pop("HORAS_CACHE_AIS_VIGENTE", None)
        shutil.rmtree(tmp, ignore_errors=True)


def main(argv: list[str] | None = None) -> None:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--sin-lake" in args:                                        # paso 5: variante sin Lake (la heredan los subprocesos)
        args.remove("--sin-lake")
        os.environ["TORRE_USAR_LAKE"] = "0"
        C.USAR_LAKE = False
    if not args:                                                    # corrida directa: manda lo de arriba
        args = [MODO] + (["--sin-lake"] if SIN_LAKE else [])
        if "--sin-lake" in args:
            args.remove("--sin-lake"); os.environ["TORRE_USAR_LAKE"] = "0"; C.USAR_LAKE = False
    modo = args[0].lower()

    if modo == "paso":
        if len(args) < 2:
            raise SystemExit("Uso: python torre.py paso <in_transit|eventos|pedidos|correo|radar|correo_radar|backfill|chequeo>")
        ejecutar_paso(args[1].lower())
        if C.AVISOS_PASO:                                           # paso 4: terminado con aviso
            print("AVISO -> " + " · ".join(C.AVISOS_PASO), flush=True)
            sys.exit(C.SALIDA_CON_AVISO)
        return

    if modo == "diario":
        main_diario_subprocesos()
        return

    if modo == "semanal":                                           # paso 5: lunes (tarea programada)
        resumen = []
        for paso in ("backfill", "foto_semanal"):                   # rescate desde los Inbound y despues la foto
            t0 = time.perf_counter()
            error = correr_subproceso(paso)
            resumen.append((paso, time.perf_counter() - t0, error))
        for paso, seg, error in resumen:
            print(f"   {paso:<13} {seg:>6.1f} s   {'OK' if error is None else 'FALLO -> ' + error}")
        if any(e and not e.startswith("AVISO ") for _, _, e in resumen):
            sys.exit(1)
        return

    if modo == "radar_horario":                                     # paso 5: tarea programada cada hora
        t0 = time.perf_counter()
        error = correr_subproceso("radar_horario")
        print(f"\n   radar_horario {time.perf_counter() - t0:>6.1f} s   {'OK' if error is None else 'FALLO -> ' + error}")
        if error and not error.startswith("AVISO "):
            sys.exit(1)
        return

    if modo in ("backfill", "radar"):
        t0 = time.perf_counter()
        error = correr_subproceso(modo)
        dt = time.perf_counter() - t0
        print(f"\n   {modo:<13} {dt:>6.1f} s   {'OK' if error is None else 'FALLO -> ' + error}")
        if error:
            sys.exit(1)
        return

    if modo == "chequeo":
        main_chequeo()
        return

    raise SystemExit(f"Modo desconocido: {modo!r} — usar: diario | backfill | radar | chequeo | paso <X>")


if __name__ == "__main__":
    main()
