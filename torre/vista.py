# -*- coding: utf-8 -*-
"""
vista.py — Construcción de las tablas y estructuras de negocio para el correo de arribados y el radar:
  - Correo y pestaña aérea del radar: inicio_en_ruta, lineas_que_salieron, pedidos_para_salidos, armar()
  - Radar marítimo y aéreo: leer_maritimos, lineas_con_pedidos, lineas_por_422, armar_bls, agrupar_buques,
    aplicar_estado, separar_vista, aereos_en_terminal(armado=None)

Redundancia R1 eliminada: armar() se calcula UNA sola vez por corrida en salidas.main_correo() y su resultado
(tabla, corte, lineas) se pasa en memoria a aereos_en_terminal(armado=armado) sin escribir archivos nuevos.
Si el radar corre suelto (python torre.py radar), aereos_en_terminal() invoca armar() directamente.
"""

import logging
from datetime import datetime, timedelta  # noqa: F401 (usado por reloj virtual en pruebas)

import numpy as np
import pandas as pd

import eventos as E  # noqa: F401 (respeta jerarquía acíclica vista -> ingesta/eventos)
import ingesta as I
import torre_config as C

# ══════════════════════════ CONFIG · VISTA CORREO & RADAR ════════════════════
COLS_RECORRIDO  = ["llegada_fecha_fuente", "atd_correo", "atd"]   # de seguimiento (eventos)
SOLO_CON_PEDIDO = False            # False: el transito manda y el pedido es visibilidad; el embarque arribado sin
                                   # ningun pedido entra con una fila de su stock. True: solo embarques con pedido B2B

HITO_TRANSITO   = "En tránsito internacional"
HITOS_PUERTO    = ("Arribado", "En tránsito a DC")      # ya en el Callao: sin consulta web
DIAS_PROXIMO    = 3                                     # disparador del 17/11
RADIO_CERCA_KM  = 1500                                  # el «Próximos» de tu radar
ESTADOS_VISTA   = ("vencida", "en_callao", "por_arribar", "en_puerto")   # paso 5: tambien lo arribado sin salida (pedido 25/11)
CALLAO          = {"locode": "PECLL", "nombre": "Callao", "lat": -12.0464, "lon": -77.1428}
HORAS_AIS_VALIDA = 48                                   # más vieja que esto: se muestra con su antigüedad, pero ya no decide el estado

COLS_REQUERIDAS = ["key", "so", "mtm", "transit_qty", "bl_awb_num", "mot", "hito", "name_of_ocean_vessel",
                   "equipment_num", "eta", "revised_eta", "outlook_eta", "llegada_aviso_mail"]
COLS_OPCIONALES = ["eta_correo", "container_size_code", "carrier_name", "site", "route", "puerto_descarga"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
log_correo = logging.getLogger("mail_arribados")
log_radar  = logging.getLogger("radar_embarques")


# ══════════════════════════ HELPERS DE FORMATO JSON / FECHAS ═════════════════
def iso(valor):
    return None if pd.isna(valor) else pd.Timestamp(valor).strftime("%Y-%m-%d")


def dias(valor, hoy):
    return None if pd.isna(valor) else int((pd.Timestamp(valor).normalize() - hoy).days)


def limpio(valor):
    """Valor de pandas a JSON: NA/NaN -> None, numpy -> nativo."""
    if valor is None or (not isinstance(valor, (list, dict)) and pd.isna(valor)):
        return None
    return valor.item() if hasattr(valor, "item") else valor


# ══════════════════════════ TABLA DE ARRIBADOS Y EN RUTA (armar) ═════════════
def inicio_en_ruta() -> pd.Timestamp:
    """Primera fecha de salida a DC que se muestra. Tope por si el in-transit tarda en soltar lo ya ingresado."""
    return pd.Timestamp.now().normalize() - pd.offsets.BDay(C.DIAS_HABILES_ENTREGA)


def _resumen(valores: pd.Series) -> str:
    """'4224382978' o '4224382978 +2' en la fila de stock: el detalle por linea va en el Excel adjunto."""
    u = sorted(valores.dropna().astype(str).unique())
    return "—" if not u else u[0] if len(u) == 1 else f"{u[0]} +{len(u) - 1}"


def lineas_que_salieron(seg: pd.DataFrame, corte) -> pd.DataFrame:
    """Lineas que Transito ya borro: aereos con aviso o salida a DC registrada, maritimos con salida a DC. Sus lineas
    salen del historial de job_in_transit; su aviso y su salida FWD, del historial de eventos. Mismas columnas que
    seguimiento: el resto del correo las trata igual (fuera_in_transit las marca)."""
    ruta = C.RUTA_HISTORIAL
    if not ruta.exists():
        log_correo.warning("sin %s: el correo solo ve lo que sigue en el in-transit (correr job_in_transit actualizado)",
                           ruta.name)
        return pd.DataFrame()
    hist = pd.read_parquet(ruta)
    ultima = hist.sort_values(["visto_desde", "visto_hasta"]).groupby("key").tail(1)       # version vigente por linea
    maritimo = ultima["mot"].astype("string").str.upper().str.strip().isin(C.MOT_MARITIMO)
    fuera = ultima.assign(_mar=maritimo.fillna(False).astype(bool))[
        ~ultima["key"].astype(str).isin(set(seg["key"].astype(str)))
        & pd.to_datetime(ultima["visto_hasta"]).ge(pd.Timestamp.now() - pd.Timedelta(days=C.DIAS_VENTANA_SALIDOS))]
    if fuera.empty or not C.RUTA_EVENTOS.exists():
        return pd.DataFrame()
    with pd.ExcelFile(C.RUTA_EVENTOS) as libro:
        hojas = {h: pd.read_excel(libro, h, dtype={"hawb_key": str, "hora_estimada": str}) for h in ("arribos", "salidas_dc")}
    ultimo = lambda df: (df.dropna(subset=["hawb_key"]).sort_values("fecha_mail")
                           .drop_duplicates("hawb_key", keep="last").set_index("hawb_key"))
    aviso, salida = ultimo(hojas["arribos"]), ultimo(hojas["salidas_dc"])
    fuera = fuera.assign(_k=C.clave_hawb(fuera["bl_awb_num"]))
    con_aviso = fuera["_k"].isin(aviso.index) & ~fuera["_mar"]            # maritimo: solo con salida a DC
    fuera = fuera[con_aviso | fuera["_k"].isin(salida.index)]
    if fuera.empty:
        return pd.DataFrame()
    a = aviso.reindex(fuera["_k"].where(~fuera["_mar"]))                  # alineados con 'fuera' por posicion;
    s = salida.reindex(fuera["_k"])                                       # el maritimo no toma nada de su aviso
    del_pdf = pd.to_datetime(a["fecha_llegada"], errors="coerce").to_numpy()
    mail_aviso = pd.to_datetime(a["fecha_mail"], errors="coerce")
    con_salida = s["fecha_mail"].notna().to_numpy()
    salidos = pd.DataFrame({
        "key": fuera["key"].to_numpy(), "so": fuera["so"].to_numpy(), "mtm": fuera["mtm"].to_numpy(),
        "transit_qty": fuera["transit_qty"].to_numpy(), "bl_awb_num": fuera["bl_awb_num"].to_numpy(),
        "mot": fuera["mot"].to_numpy(), "delivery_order_number": fuera["delivery_order_number"].to_numpy(),
        "hito": np.where(con_salida, "En tránsito a DC", C.ARRIBADO),
        "llegada_fecha": pd.Series(del_pdf).fillna(pd.Series(mail_aviso.dt.normalize().to_numpy())).to_numpy(),
        "llegada_fecha_fuente": np.where(pd.notna(del_pdf), "PDF del aviso",
                                         np.where(mail_aviso.notna().to_numpy(), "fecha del correo", None)),
        "llegada_aviso_mail": mail_aviso.to_numpy(),
        "salida_dc_fecha": pd.to_datetime(s["fecha_programada"], errors="coerce").to_numpy(),
        "salida_dc_hora": s["hora_estimada"].to_numpy(),
        "salida_dc_mail": pd.to_datetime(s["fecha_mail"], errors="coerce").to_numpy(),
        "corte_bi": corte, "fuera_in_transit": True,
        **{c: pd.to_numeric(a[c.replace("aviso_", "bultos_")], errors="coerce").to_numpy()
           for c in C.COLS_AVISO if c.replace("aviso_", "bultos_") in a.columns}})
    log_correo.info("fuera del in-transit con aviso o salida a DC (historial): %d lineas · %d HAWB/BL · por mot %s",
                    len(salidos), salidos["bl_awb_num"].nunique(),
                    salidos.drop_duplicates("bl_awb_num")["mot"].value_counts().to_dict())
    return salidos


def pedidos_para_salidos(pedidos_hoy: pd.DataFrame) -> pd.DataFrame:
    """Pedidos para las lineas que ya salieron del in-transit: el OrderMatch de hoy ya no trae su 433 (su match de
    transito se fue con la linea), asi que lo que hoy no esta sale del ultimo visto en historial_pedidos.parquet (copia
    diaria de job_in_transit; para lo que salio antes, el Order Match del Inbound semanal). Lo de hoy manda."""
    ruta = C.RUTA_HISTORIAL_PEDIDOS
    if not ruta.exists():
        log_correo.warning("sin %s: lo que salio del in-transit solo cruza con el OrderMatch de hoy (correr job_in_transit "
                           "actualizado y el backfill)", ruta.name)
        return pedidos_hoy
    previos = pd.read_parquet(ruta, columns=["so", "mtm", "so_local", "cliente"])
    previos = previos.assign(so=C.txt(previos["so"]), mtm=C.txt(previos["mtm"]).str.upper(),
                             so_local=C.txt(previos["so_local"]), cliente=previos["cliente"].astype("string"))
    par = lambda d: d["so"].astype(str) + "|" + d["mtm"].astype(str)
    extra = previos[~par(previos).isin(set(par(pedidos_hoy)))]
    log_correo.info("pedidos de lo que salio del in-transit: %d 422 x MTM del historial (el OrderMatch de hoy ya no los trae)",
                    extra[["so", "mtm"]].drop_duplicates().shape[0])
    return pd.concat([pedidos_hoy, extra], ignore_index=True)


def sla_de_mot(mot) -> float:
    """SLA en dias habiles de un embarque. Si junta varios MOT ("AIR / SEA") toma el mas estricto, para que
    nunca quede sin SLA (antes daba NaN y el embarque no salia fuera de SLA ni en el radar aereo)."""
    slas = [C.SLA_ARRIBO_HAB[m] for m in str(mot).split(" / ") if m in C.SLA_ARRIBO_HAB]
    return min(slas) if slas else float("nan")


def armar():
    """Filas: embarque x 422 x MTM con pedido B2B (cantidad del transito, lista de (433, cliente)) y el estado del
    embarque (Arribado / En ruta al DC).
    Por embarque, en crudo: unidades totales del HAWB (B2B + stock / e-commerce) y bultos, kg, m3 de Lake; y para el
    recorrido y las alertas: ATD, llegada y su fuente, dias habiles desde la llegada, SLA y conteo del aviso."""
    seg, _, corte = I.leer_seguimiento()
    salidos = lineas_que_salieron(seg, corte)       # aereos que Transito ya borro: llegaron al terminal
    seg = seg.assign(fuera_in_transit=False)
    if len(salidos):
        seg = pd.concat([seg, salidos], ignore_index=True)
    pedidos = I.pedidos_de_om(I.leer_ordermatch())
    fuera = seg["fuera_in_transit"].fillna(False).astype(bool)
    largo = I.lineas_con_pedido(seg[~fuera], pedidos)          # en el bi: el OrderMatch de hoy
    if fuera.any():                                            # fuera: hoy, y lo que hoy ya no trae, del historial
        largo = pd.concat([largo, I.lineas_con_pedido(seg[fuera], pedidos_para_salidos(pedidos))], ignore_index=True)
    for col in COLS_RECORRIDO:                    # sin ellas el correo sale igual, con "—" en el recorrido
        if col not in largo.columns:
            log_correo.warning("seguimiento sin %r: correr el job_eventos_transito actual", col)
            largo[col] = pd.NA
    desde_ruta = inicio_en_ruta()
    salida = pd.to_datetime(largo["salida_dc_fecha"], errors="coerce").fillna(          # sin fecha legible en el
        pd.to_datetime(largo["salida_dc_mail"], errors="coerce").dt.normalize())        # correo FWD: la del correo
    con_hawb, con_salida = largo["bl_awb_num"].notna(), largo["salida_dc_mail"].notna()
    # hito de job_eventos_transito: aereo con aviso; maritimo con ATA o ETA vencida (su aviso de salida no es arribo)
    arribado = con_hawb & largo["hito"].eq(C.ARRIBADO) & ~con_salida
    en_ruta = con_hawb & con_salida & salida.ge(desde_ruta)
    largo = largo.assign(estado=np.select([arribado, en_ruta], [C.ARRIBADO, C.EN_RUTA], default=""), salida=salida,
                         salida_hora=C.txt(largo.get("salida_dc_hora", pd.Series(pd.NA, index=largo.index))))
    arr = largo[arribado | en_ruta]
    embarques = lambda m: largo.loc[m, "bl_awb_num"].nunique()
    log_correo.info("embarques: %d arribados sin salida · %d en ruta al DC (salida desde %s) · %d con salida anterior "
                    "(ya en DC, fuera)", embarques(arribado), embarques(en_ruta), f"{desde_ruta:%d/%m}",
                    embarques(con_hawb & con_salida & ~en_ruta))
    log_correo.info("en el correo: %d lineas de transito · %s", arr["key"].nunique(),
                    arr.drop_duplicates("key")["vinculo"].value_counts().to_dict())
    con = arr[arr["vinculo"].eq("con pedido")].assign(cliente=lambda d: d["cliente"].fillna("(sin cliente en OM)"))
    llave = ["bl_awb_num", "so", "mtm"]
    # cantidad: cada linea del bi una sola vez (en 'largo' una linea con dos pedidos aparece dos veces)
    lineas = (con.drop_duplicates("key").assign(llegada=lambda d: pd.to_datetime(d["llegada_fecha"], errors="coerce"))
                 .groupby(llave).agg(cantidad=("transit_qty", "sum"), llegada=("llegada", "max")))
    unicos = con.drop_duplicates(llave + ["so_local"])
    pedidos = (unicos.assign(pedidos=list(zip(unicos["so_local"], unicos["cliente"])))
                     .groupby(llave)["pedidos"].agg(list))                      # [(433, cliente), ...] por fila
    # el embarque completo, en crudo: todas sus lineas del bi (tambien stock / e-commerce) y sus bultos de Lake
    todas = arr.drop_duplicates("key")
    fecha = lambda col: pd.to_datetime(todas[col], errors="coerce")
    todas = todas.assign(_llegada=fecha("llegada_fecha"),
                         _atd=fecha("atd_correo").fillna(fecha("atd")))          # ATD: correos FWD, si no el reporte
    por_embarque = todas.groupby("bl_awb_num").agg(
        mot=("mot", lambda x: " / ".join(sorted(x.dropna().astype(str).str.upper().unique()))),
        unidades_embarque=("transit_qty", "sum"),
        estado=("estado", "first"), salida=("salida", "max"), salida_hora=("salida_hora", "first"),
        llegada_embarque=("_llegada", "max"), llegada_fuente=("llegada_fecha_fuente", "first"),
        atd_origen=("_atd", "max"), fuera_in_transit=("fuera_in_transit", "max"))         # por HAWB
    por_embarque["fuera_in_transit"] = por_embarque["fuera_in_transit"].fillna(False).astype(bool)
    por_embarque["dias_hab"] = C.dias_habiles_no_neg(por_embarque["llegada_embarque"], pd.Timestamp.now().normalize())
    por_embarque["sla_hab"] = por_embarque["mot"].map(sla_de_mot)          # doble MOT: el SLA mas estricto (paso 4)
    por_embarque["fuera_sla"] = por_embarque["estado"].eq(C.ARRIBADO) & por_embarque["dias_hab"].gt(por_embarque["sla_hab"])
    por_embarque = por_embarque.join(C.leer_embarques(C.COLS_EMBARQUE, deduplicar=False, logger=log_correo).rename(columns={
        "n_bultos": "bultos_embarque", "peso_bruto_kg": "kg_embarque", "volumen_m3": "m3_embarque"}))
    # fuera del in-transit no estan en la hoja embarques: el conteo y los bultos salen de su aviso
    del_aviso = todas.groupby("bl_awb_num")[[c for c in C.COLS_AVISO if c in todas.columns]].first()
    for col in del_aviso.columns:
        por_embarque[col] = por_embarque[col].fillna(del_aviso[col])
    por_embarque["bultos_embarque"] = por_embarque["bultos_embarque"].fillna(
        por_embarque["aviso_total"].where(por_embarque["fuera_in_transit"]))
    aviso = lambda col: pd.to_numeric(por_embarque[col], errors="coerce").fillna(0)
    por_embarque["alerta_aviso"] = aviso("aviso_malos").gt(0) | aviso("aviso_faltantes").gt(0)
    filas = lineas.join(pedidos)
    sin_pedido = todas[~todas["bl_awb_num"].isin(set(con["bl_awb_num"]))]
    if not SOLO_CON_PEDIDO and len(sin_pedido):       # embarque arribado sin ningun pedido: una fila con su stock
        stock = (sin_pedido.groupby("bl_awb_num")                      # funcion en agg: pocos embarques
                 .agg(so=("so", _resumen), mtm=("mtm", _resumen), unid_stock=("transit_qty", "sum"),
                      llegada=("_llegada", "max"))
                 .assign(cantidad=0.0).reset_index().set_index(llave))
        stock["pedidos"] = [[] for _ in range(len(stock))]
        filas = pd.concat([filas, stock])
        log_correo.info("embarques arribados sin ningun pedido (entran con su stock): %d", len(stock))
    tabla = filas.join(por_embarque, on="bl_awb_num").reset_index().rename(columns={"bl_awb_num": "embarque"})
    sold = C.sold_to_por_pedido(avisar_sin_columna=True, logger=log_correo)
    tabla["sold_to"] = [[sold.get(p) for p, _ in lista] for lista in tabla["pedidos"]]   # alineado con (433, cliente)
    log_correo.info("sold-to: %d de %d pedidos del correo", sum(s is not None for l in tabla["sold_to"] for s in l),
                    sum(len(l) for l in tabla["pedidos"]))
    # arribados por llegada, en ruta por salida
    tabla = (tabla.assign(_e=tabla["estado"].map({C.ARRIBADO: 0, C.EN_RUTA: 1}),
                          _f=tabla["llegada"].where(tabla["estado"].eq(C.ARRIBADO), tabla["salida"]))
                  .sort_values(["_e", "_f", "embarque", "so", "mtm"], na_position="last")
                  .drop(columns=["_e", "_f"]).reset_index(drop=True))
    emb = tabla.drop_duplicates("embarque")
    log_correo.info("requieren accion: %d de %d embarques · %d fuera de SLA · %d con bultos malos o faltantes en el aviso",
                    int((emb["fuera_sla"] | emb["alerta_aviso"]).sum()), len(emb), int(emb["fuera_sla"].sum()),
                    int(emb["alerta_aviso"].sum()))
    # adjunto: TODAS las lineas de los embarques del correo (B2B + stock), una fila por linea del bi (se puede sumar)
    lineas = I.por_linea(arr[arr["bl_awb_num"].isin(set(tabla["embarque"]))])
    return tabla, corte, lineas


# ══════════════════════════ RADAR · LECTURA Y ÁRBOL MARÍTIMO ═════════════════
def maritimos_fuera_bi(seg: pd.DataFrame) -> pd.DataFrame:
    """Paso 5 (ESC-06): marítimos que Transito ya soltó pero siguen en el terminal sin salida a DC. Salen del historial
    (última versión, con su buque desde el paso 5) y su arribo sigue la regla 4: ATA del correo FWD o ETA vencida.
    Lo que ya tiene salida a DC no entra aquí: va a la vista Salida a DC. Mismas columnas que seguimiento."""
    if not C.RUTA_HISTORIAL.exists():
        return pd.DataFrame()
    hist = I.leer_historial_transito()
    ultima = hist.sort_values(["visto_desde", "visto_hasta"]).groupby("key").tail(1)
    mar = ultima["mot"].astype("string").str.upper().str.strip().isin(C.MOT_MARITIMO).fillna(False).astype(bool)
    reciente = pd.to_datetime(ultima["visto_hasta"]).ge(pd.Timestamp.now() - pd.Timedelta(days=C.DIAS_VENTANA_SALIDOS))
    fuera = ultima[mar & reciente & ~ultima["key"].astype(str).isin(set(seg["key"].astype(str)))].copy()
    if fuera.empty:
        return pd.DataFrame()
    fuera["_k"] = C.clave_hawb(fuera["bl_awb_num"])
    con_salida, ata = set(), pd.Series(dtype="datetime64[ns]")
    if C.RUTA_EVENTOS.exists():
        with pd.ExcelFile(C.RUTA_EVENTOS) as libro:
            salidas = pd.read_excel(libro, "salidas_dc", dtype={"hawb_key": str})
            notif = pd.read_excel(libro, "notificaciones", dtype={"hawb_key": str}) if "notificaciones" in libro.sheet_names else None
        con_salida = set(salidas["hawb_key"].dropna())
        if notif is not None and "ata" in notif:
            ata = pd.to_datetime(notif["ata"], errors="coerce").groupby(notif["hawb_key"]).max()
    fuera = fuera[~fuera["_k"].isin(con_salida)]
    if fuera.empty:
        return pd.DataFrame()
    llegada = fuera["_k"].map(ata)
    eta = pd.to_datetime(fuera["revised_eta"], errors="coerce").fillna(pd.to_datetime(fuera["outlook_eta"], errors="coerce"))
    arribado = llegada.notna() | eta.le(pd.Timestamp.now().normalize())
    extra = fuera.reindex(columns=COLS_REQUERIDAS).assign(
        hito=np.where(arribado, C.ARRIBADO, HITO_TRANSITO), eta=eta, eta_correo=llegada, llegada_aviso_mail=pd.NaT)
    log_radar.info("marítimo fuera del in-transit sin salida a DC (historial): %d líneas · %d BL · %d arribados",
                   len(extra), extra["bl_awb_num"].nunique(), int(arribado.sum()))
    return extra


def alertas_eta(hoy=None) -> list:
    """Paso 5: BL cuya ETA vigente (revised_eta, si no outlook_eta) se movió más de C.UMBRAL_ALERTA_ETA_DIAS contra la
    versión anterior del historial (desde el paso 5 un cambio de ETA abre versión). Regla de Hi: el compromiso se mide
    sobre la ETA más reciente, así que la alerta compara la nueva contra la anterior. Solo lo visto en la ventana."""
    if not C.RUTA_HISTORIAL.exists():
        return []
    hoy = pd.Timestamp.now() if hoy is None else hoy
    h = I.leer_historial_transito()
    h = h.assign(_eta=pd.to_datetime(h["revised_eta"], errors="coerce").fillna(pd.to_datetime(h["outlook_eta"], errors="coerce")))
    h = h[pd.to_datetime(h["visto_hasta"]).ge(hoy - pd.Timedelta(days=C.DIAS_VENTANA_SALIDOS))]
    porbl = (h.dropna(subset=["bl_awb_num", "_eta"]).groupby(["bl_awb_num", "visto_desde"], as_index=False)
              .agg(_eta=("_eta", "max"), buque=("name_of_ocean_vessel", "first"), mot=("mot", "first"))
              .sort_values(["bl_awb_num", "visto_desde"]))
    alertas = []
    for bl, g in porbl.groupby("bl_awb_num"):
        if len(g) < 2:
            continue
        ant, nue = g.iloc[-2], g.iloc[-1]
        delta = (nue["_eta"] - ant["_eta"]).days
        if abs(delta) > C.UMBRAL_ALERTA_ETA_DIAS:
            alertas.append({"bl": str(bl), "buque": limpio(nue["buque"]), "mot": limpio(nue["mot"]),
                            "eta_anterior": iso(ant["_eta"]), "eta_nueva": iso(nue["_eta"]), "delta_dias": int(delta),
                            "visto": iso(nue["visto_desde"])})
    log_radar.info("alertas de ETA (> %d días): %d BL", C.UMBRAL_ALERTA_ETA_DIAS, len(alertas))
    return sorted(alertas, key=lambda a: -abs(a["delta_dias"]))


def dn_por_433() -> dict:
    """Paso 5: 433 -> DN local del M4 (con y sin GI). Si la 433 ya tiene DN, SAP ya tenía el stock: «ingresado en SAP»
    (misma prueba que diag_ingreso_dc). Sin M4 legible la vista sale igual, sin el estado SAP."""
    partes = []
    for nombre in C.M4_ARCHIVOS:
        ruta = C.CARPETA_M4 / nombre
        if ruta.exists():
            m = pd.read_parquet(ruta, columns=["site_country", "so", "dn"])
            partes.append(m[m["site_country"].astype("string").str.upper().eq(C.PAIS_M4)])
    if not partes:
        log_radar.warning("sin M4 en %s: la vista Salida a DC sale sin estado SAP", C.CARPETA_M4)
        return {}
    m4 = pd.concat(partes, ignore_index=True).dropna(subset=["dn"])
    return dict(zip(C.txt(m4["so"]), C.txt(m4["dn"])))


def salida_dc(hoy=None) -> list:
    """Paso 5: vista Salida a DC por embarque: programada / en ruta al DC / entregado al DC (regla 6), con sus 433 y si
    cada una ya está ingresada en SAP (DN en el M4) o pendiente. Lee seguimiento + lo que Transito ya soltó."""
    hoy = (pd.Timestamp.now() if hoy is None else hoy).normalize()
    seg, _, _ = I.leer_seguimiento()
    fuera = lineas_que_salieron(seg, seg["corte_bi"].iloc[0] if len(seg) else hoy)
    lin = pd.concat([seg, fuera], ignore_index=True) if len(fuera) else seg
    lin = lin[lin["hito"].isin(["En tránsito a DC", "Entregado al DC"])]
    if lin.empty:
        return []
    try:
        pedidos = pedidos_para_salidos(I.pedidos_de_om(I.leer_ordermatch()))
    except (AssertionError, RuntimeError, OSError, ValueError) as e:
        log_radar.warning("salida a DC sin pedidos (433): %s", e)
        pedidos = pd.DataFrame(columns=["so", "mtm", "so_local", "cliente"], dtype="string")
    largo = I.lineas_con_pedido(lin, pedidos)
    dn = dn_por_433()
    salida = []
    for emb, g in largo.groupby("bl_awb_num"):
        g0 = g.iloc[0]
        fecha = pd.to_datetime(g["salida_dc_fecha"], errors="coerce").max()
        estado = ("Entregado al DC" if g0["hito"] == "Entregado al DC"
                  else "Programada" if pd.notna(fecha) and fecha > hoy else "En ruta al DC")
        peds = []
        for p, c in g[["so_local", "cliente"]].dropna(subset=["so_local"]).drop_duplicates().itertuples(index=False):
            p = str(p)
            peds.append({"so433": p, "cliente": limpio(c), "dn": dn.get(p),
                         "sap": ("sin M4" if not dn else "ingresado" if p in dn else "pendiente")})
        salida.append({"embarque": str(emb), "mot": limpio(g0["mot"]), "estado": estado, "salida_dc": iso(fecha),
                       "recepcion_dc": iso(pd.to_datetime(g.get("recepcion_dc"), errors="coerce").max()) if "recepcion_dc" in g else None,
                       "unidades": float(pd.to_numeric(g.drop_duplicates("key")["transit_qty"], errors="coerce").sum()),
                       "pedidos": peds, "pendientes_sap": sum(x["sap"] == "pendiente" for x in peds)})
    log_radar.info("salida a DC: %d embarques · %d 433 pendientes de SAP", len(salida), sum(s["pendientes_sap"] for s in salida))
    return sorted(salida, key=lambda s: ({"Entregado al DC": 0, "En ruta al DC": 1, "Programada": 2}[s["estado"]], -s["pendientes_sap"]))


def leer_maritimos():
    seg, _, corte = I.leer_seguimiento()                 # aborta si seguimiento tiene más de 12 h
    extra = maritimos_fuera_bi(seg)                      # paso 5: lo que Transito soltó y sigue en el terminal
    if len(extra):
        seg = pd.concat([seg, extra], ignore_index=True)
    faltan = [c for c in COLS_REQUERIDAS if c not in seg.columns]
    assert not faltan, f"seguimiento no trae {faltan} — revisar job_in_transit / job_eventos_transito"
    seg = seg.reindex(columns=list(dict.fromkeys([*seg.columns, *COLS_OPCIONALES])))
    mot = seg["mot"].astype("string").str.strip().str.upper()
    hito = seg["hito"].astype("string").str.strip()
    bl = C.txt(seg["bl_awb_num"])
    maritimo = mot.isin(C.MOT_MARITIMO).fillna(False).astype(bool)
    vigente = hito.isin([HITO_TRANSITO, *HITOS_PUERTO]).fillna(False).astype(bool)
    fuera = {"aereas": int((~maritimo).sum()), "sin_bl": int((maritimo & bl.isna()).sum()),
             "entregadas": int((maritimo & bl.notna() & ~vigente).sum())}
    dentro = maritimo & bl.notna() & vigente
    # columnas ya filtradas: asignar una Series larga a un DataFrame vacío le copia su índice (filas fantasma)
    mar = seg[dentro].assign(
        mot=mot[dentro], hito=hito[dentro], bl=bl[dentro], contenedor=C.txt(seg.loc[dentro, "equipment_num"]),
        buque=C.txt(seg.loc[dentro, "name_of_ocean_vessel"]).str.upper().str.replace(r"\s+", " ", regex=True))
    for col in ("eta", "revised_eta", "outlook_eta", "eta_correo", "llegada_aviso_mail"):
        mar[col] = pd.to_datetime(mar[col], errors="coerce").dt.normalize()
    return mar, fuera, corte


def lineas_con_pedidos(mar: pd.DataFrame):
    """Cruce de job_pedidos_transito. Sin OrderMatch legible el radar sale igual: el pedido es visibilidad."""
    try:
        pedidos, aviso = I.pedidos_de_om(I.leer_ordermatch()), None
        pedidos = pedidos_para_salidos(pedidos)       # paso 5: lo que Transito soltó toma su 433 del historial (como el correo)
    except (AssertionError, RuntimeError, OSError, ValueError) as e:
        log_radar.warning("radar sin pedidos (433): %s", e)
        pedidos, aviso = pd.DataFrame(columns=["so", "mtm", "so_local", "cliente"], dtype="string"), str(e)
    largo = I.lineas_con_pedido(mar, pedidos)
    assert largo["key"].nunique() == len(mar), "el cruce con OrderMatch perdió líneas del tránsito"
    return largo, aviso


def lineas_por_422(f: pd.DataFrame, pedidos_de) -> list:
    """422 x MTM del embarque (el grano del cruce y del correo): cantidad con cada línea del bi una sola vez, su vínculo
    y sus 433 (pedidos_de(so, mtm) -> lista). Mismo formato en las dos pestañas: la rama de la plantilla es una."""
    unicas = f.drop_duplicates("key").assign(so=lambda d: d["so"].fillna("—"), mtm=lambda d: d["mtm"].fillna("—"))
    return [{"so": so, "mtm": mtm, "cantidad": limpio(g["transit_qty"].sum(min_count=1)),
             "vinculo": limpio(g["vinculo"].iloc[0]), "pedidos": pedidos_de(so, mtm)}
            for (so, mtm), g in unicas.groupby(["so", "mtm"], sort=False)]


def armar_bls(largo: pd.DataFrame, sold: dict, emb: pd.DataFrame, hoy) -> list:
    """Una entrada por BL con fechas, carga y líneas. Bucle por BL: son decenas, y el anidado
    (BL -> 422 · MTM -> pedidos) es lo que lee la plantilla."""
    largo = largo.assign(sold_to=largo["so_local"].map(sold))
    unicos = lambda s: [str(v) for v in dict.fromkeys(s.dropna())]
    bls = []
    for bl, f in largo.groupby("bl", sort=False):
        por_linea = f.drop_duplicates("key")
        nombres = por_linea["buque"].dropna()
        buque = nombres.mode().iloc[0] if len(nombres) else None        # transbordo o dato mezclado: el de más líneas
        eta_fwd, outlook = f["eta_correo"].min(), f["outlook_eta"].min()
        eta_rep = f["revised_eta"].fillna(f["eta"]).min()
        desde = min([d for d in (eta_fwd, eta_rep) if pd.notna(d)], default=outlook)
        tope = outlook if pd.notna(outlook) else eta_rep
        hasta = max([d for d in (tope, eta_fwd) if pd.notna(d)], default=pd.NaT)
        pares = {}                                # (422, MTM) -> sus 433, sin repetir
        for p in f.dropna(subset=["so_local"]).drop_duplicates(["so", "mtm", "so_local"]).itertuples():
            pares.setdefault((p.so, p.mtm), []).append(
                {"so433": limpio(p.so_local), "cliente": limpio(p.cliente), "sold_to": limpio(p.sold_to)})
        lineas = lineas_por_422(f, lambda so, mtm: pares.get((so, mtm), []))
        carga = emb.loc[bl] if bl in emb.index else pd.Series(dtype=float)
        bls.append({
            "bl": bl, "buque": buque, "otros_buques": [n for n in unicos(nombres) if n != buque],
            "hito": limpio(f["hito"].iloc[0]), "contenedores": unicos(f["contenedor"]),
            "tipo_contenedor": unicos(f["container_size_code"].astype("string")),
            "naviera": next(iter(unicos(f["carrier_name"])), None), "ruta": next(iter(unicos(f["route"])), None),
            "origen": next(iter(unicos(f["site"])), None), "puerto": next(iter(unicos(f["puerto_descarga"])), None),
            "eta_fwd": iso(eta_fwd), "eta_reporte": iso(eta_rep), "outlook": iso(outlook),
            "desde": iso(desde), "hasta": iso(hasta), "dias_desde": dias(desde, hoy), "dias_hasta": dias(hasta, hoy),
            "aviso": iso(f["llegada_aviso_mail"].max()), "aviso_estado": None, "aviso_base": None,
            "unidades": limpio(por_linea["transit_qty"].sum(min_count=1)),
            "bultos": limpio(carga.get("n_bultos")), "kg": limpio(carga.get("peso_bruto_kg")),
            "m3": limpio(carga.get("volumen_m3")), "lineas": lineas})
    return bls


def agrupar_buques(bls: list) -> list:
    """Un buque por nombre; un BL sin nombre de buque en el reporte va solo (sin consulta web)."""
    grupos = {}
    for b in bls:
        grupos.setdefault(b["buque"] or f"(sin buque) {b['bl']}", []).append(b)
    buques = []
    for i, lista in enumerate(grupos.values()):
        transito = [b for b in lista if b["hito"] == HITO_TRANSITO]
        d_desde = min((b["dias_desde"] for b in transito if b["dias_desde"] is not None), default=None)
        d_hasta = min((b["dias_hasta"] for b in transito if b["dias_hasta"] is not None), default=None)
        d_hasta_max = max((b["dias_hasta"] for b in transito if b["dias_hasta"] is not None), default=None)
        if not transito:
            por_fecha = "en_puerto"
        elif d_hasta is not None and d_hasta < 0:
            por_fecha = "vencida"
        elif d_desde is None:
            por_fecha = "sin_fecha"
        else:
            por_fecha = "por_arribar" if d_desde <= DIAS_PROXIMO else "en_transito"
        buques.append({"id": f"b{i}", "nombre": lista[0]["buque"], "por_fecha": por_fecha, "dias_desde": d_desde,
                       "dias_hasta": d_hasta, "dias_hasta_max": d_hasta_max,
                       "naviera": next((b["naviera"] for b in lista if b["naviera"]), None),
                       "bls": lista})
    return buques


def aplicar_estado(buque: dict, ais: dict, ficha: dict, ahora) -> None:
    """Estado final del buque y del aviso de cada BL. El AIS decide solo si su posición tiene menos de
    HORAS_AIS_VALIDA h; si no, manda la fecha."""
    pos = (ais or {}).get("pos")
    edad = C.edad_horas((ais or {}).get("consultado"), ahora) if pos else None
    valida = pos is not None and edad <= HORAS_AIS_VALIDA
    cerca = valida and (pos["arribado"] or pos["dist_km"] <= RADIO_CERCA_KM)
    if buque["por_fecha"] == "en_puerto":
        estado = "en_puerto"
    elif valida and pos["arribado"]:
        estado = "en_callao"
    elif buque["por_fecha"] == "vencida":
        estado = "vencida"
    elif cerca or buque["por_fecha"] == "por_arribar":
        estado = "por_arribar"
    else:
        estado = "en_transito"
    for bl in buque["bls"]:
        if bl["hito"] == HITO_TRANSITO and bl["aviso"]:
            cerca_bl = cerca if valida else (bl["dias_desde"] is not None and bl["dias_desde"] <= DIAS_PROXIMO)
            bl["aviso_estado"], bl["aviso_base"] = ("vigente" if cerca_bl else "anticipado"), ("AIS" if valida else "fecha")
    if ais:
        nota = ais.get("nota")
    elif buque["por_fecha"] == "en_puerto":
        nota = "Ya en el Callao según la cadena: no se consulta"
    elif not buque["nombre"]:
        nota = "El reporte no trae el nombre del buque"
    else:
        nota = f"No se consulta: llega en más de {DIAS_PROXIMO} días"
    buque.update(estado=estado, ais=pos, ais_origen=(ais or {}).get("origen") if pos else None,
                 ais_consultado=(ais or {}).get("consultado") if pos else None,
                 ais_edad_h=round(edad, 1) if pos else None, ais_valida=valida, ais_nota=nota,
                 imo=ficha.get("imo"), mmsi=ficha.get("mmsi"),
                 nombre_dudoso=ficha.get("exacto") is False)


def separar_vista(buques: list) -> tuple:
    """Pestaña marítima: buques con nombre y cerca (ESTADOS_VISTA). Lo demás no se dibuja: se cuenta por motivo."""
    en_vista = lambda b: bool(b["nombre"]) and b["estado"] in ESTADOS_VISTA
    fuera = {"lejos": 0, "sin_fecha": 0, "sin_nombre_bl": 0, "arribados_bl": 0, "avisos_anticipados": 0}
    for b in buques:
        if en_vista(b):
            continue
        if b["estado"] == "en_puerto":
            fuera["arribados_bl"] += len(b["bls"])
        elif not b["nombre"]:
            fuera["sin_nombre_bl"] += len(b["bls"])
        elif b["dias_desde"] is None and not b["ais_valida"]:
            fuera["sin_fecha"] += 1
        else:
            fuera["lejos"] += 1
        fuera["avisos_anticipados"] += sum(x["aviso_estado"] == "anticipado" for x in b["bls"])
    return [b for b in buques if en_vista(b)], fuera


# ══════════════════════════ RADAR · AÉREOS EN EL TERMINAL ════════════════════
def aereos_en_terminal(armado=None) -> dict:
    """HAWB aéreos con aviso de llegada y sin salida a DC: filas «Arribado» de armar(). Misma regla,
    mismo cruce (OrderMatch de hoy + historial para lo que Transito ya borró) y el conteo físico de la extracción
    de los avisos. Si `armado` viene ya calculado en memoria desde el paso correo (R1), se reutiliza sin recalcular;
    si es None (corrida suelta del radar), invoca armar(). Si armar() falla, el marítimo sale igual."""
    salida = {"hawbs": [], "en_ruta": 0, "sla_hab": C.SLA_ARRIBO_HAB[C.MOT_AEREO], "error": None,
              "paso_fuera_sla": C.PASO_FUERA_SLA, "paso_aviso": C.PASO_AVISO}    # el «siguiente paso» del correo
    if armado is not None:
        tabla, _, lineas = armado
    else:
        nivel = log_correo.level
        log_correo.setLevel(logging.WARNING)            # su bitácora es la del correo: aquí solo sus advertencias
        try:
            tabla, _, lineas = armar()
        except (AssertionError, RuntimeError, OSError) as e:               # entrada que falta o vieja: el mensaje basta
            log_radar.warning("pestaña aérea sin datos (el marítimo sale igual): %s", e)
            return salida | {"error": str(e)}
        except (ValueError, KeyError, ImportError) as e:                   # inesperado: con traza para el log
            log_radar.exception("pestaña aérea sin datos (el marítimo sale igual)")
            return salida | {"error": f"{type(e).__name__}: {e}"}
        finally:
            log_correo.setLevel(nivel)
    aire = tabla[tabla["mot"].map(lambda m: C.MOT_AEREO in str(m).split(" / "))]   # doble MOT con AIR entra (paso 4)
    en_terminal = aire[aire["estado"].eq(C.ARRIBADO)]
    pares = {}                                 # (HAWB, 422, MTM) -> sus 433: (433, cliente) y sold-to van alineados
    for fila in en_terminal.itertuples():
        pares[(fila.embarque, fila.so, fila.mtm)] = [
            {"so433": limpio(p), "cliente": limpio(c), "sold_to": limpio(s)}
            for (p, c), s in zip(fila.pedidos, fila.sold_to)]
    lineas = lineas.assign(bl=C.txt(lineas["bl_awb_num"]))
    for hawb, f in lineas[lineas["bl"].isin(set(en_terminal["embarque"]))].groupby("bl", sort=False):
        cab = en_terminal[en_terminal["embarque"].eq(hawb)].iloc[0]       # campos del embarque: iguales en sus filas
        salida["hawbs"].append({
            "hawb": hawb, "fuera_sla": bool(cab["fuera_sla"]), "dias_hab": limpio(cab["dias_hab"]),
            "sla_hab": limpio(cab["sla_hab"]), "llegada": iso(cab["llegada_embarque"]),
            "llegada_fuente": limpio(cab["llegada_fuente"]), "atd": iso(cab["atd_origen"]),
            "fuera_in_transit": bool(cab["fuera_in_transit"]), "unidades": limpio(cab["unidades_embarque"]),
            "bultos": limpio(cab["bultos_embarque"]), "kg": limpio(cab["kg_embarque"]), "m3": limpio(cab["m3_embarque"]),
            **{col: limpio(cab.get(col)) for col in C.COLS_AVISO}, "alerta_aviso": bool(cab["alerta_aviso"]),
            "lineas": lineas_por_422(f, lambda so, mtm: pares.get((hawb, so, mtm), []))})
    salida["hawbs"].sort(key=lambda h: (not h["fuera_sla"], -(h["dias_hab"] or 0), h["llegada"] or ""))
    salida["en_ruta"] = int(aire.loc[aire["estado"].eq(C.EN_RUTA), "embarque"].nunique())
    return salida
