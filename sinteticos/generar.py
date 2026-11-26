# -*- coding: utf-8 -*-
"""
generar.py — Generador determinista de datos sintéticos y mocks para la línea base (PASO 1).

Genera únicamente datos inventados (prefijos SINT-, buques SINT AURORA / SINT PACIFICO / ...)
en sandbox/ para simular 2 días de corrida (Día 1 = ayer, Día 2 = hoy):
  - sandbox/inbound_&_OM/OM.xlsx (+ OM_dia1.xlsx y OM_dia2.xlsx para RefreshAll por día)
  - sandbox/deliverys_&_returns/m4_latam.parquet y m4_latam_sin_gi.parquet
  - sandbox/Tracking/mapa_mundial_50m.json (TopoJSON mínimo falso sin red)
  - sandbox/_sinteticos/: PDFs de avisos (texto + escaneado sin texto) y 3 pares de Excel semanales
    ("Inbound mm.dd.yyyy.xlsx" con hoja Transit tipo Copy&Paste + "Order Match mm.dd.yyyy.xlsx")

Uso: python sinteticos/generar.py [--hoy 2026-11-25] [--semilla 20261125]
"""

from __future__ import annotations

import argparse
import json
import random
import shutil
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

from fakes import AdjuntoMail, CorreoFalso, EstadoFakes, RelojVirtual, escribir_pdf_sintetico

SEMILLA_DEFAULT = 20261125
FECHA_HOY_DEFAULT = "2026-11-25"

RAIZ_V2 = Path(__file__).resolve().parent.parent
SANDBOX = RAIZ_V2 / "sandbox"
DIR_INBOUND = SANDBOX / "inbound_&_OM"
DIR_TRACKING = SANDBOX / "Tracking"
DIR_M4 = SANDBOX / "deliverys_&_returns"
DIR_INSUMOS = SANDBOX / "_sinteticos"

# Las 25 columnas exactas exigidas por SQL_HU en job_in_transit.py
HU_COLS_25 = [
    "delivery_number", "handling_unit_id", "sscc", "batch_number",
    "length", "width", "height", "unit_of_dimension",
    "total_volume", "volume_unit",
    "net_weight", "total_weight", "chargeable_weight", "volume_weight", "weight_unit",
    "tare_weight", "weight_unit_tare",
    "hu_quantity", "delivery_deleted",
    "ups_1z_number", "return_1z_number_hu",
    "gps_device_id", "gps_device_vendor_id", "gps_shipment_id",
    "loc_updated_time",
]

# Las 30 columnas Transito[...] (cubre las requeridas de job_in_transit, eventos, pedidos y radar)
TRANSITO_COLS = [
    "Key", "Delivery Order Number", "Delivery Item Number", "Transit Qty",
    "SO", "MTM", "MOT", "Country", "Storage Loc", "BL/AWB Num",
    "Ship Date", "ETA", "Revised ETA", "Outlook ETA",
    "ETD", "ATD", "ATA", "ETA Warehouse", "Pick Up", "Create Date",
    "DC STO", "PO", "Bill", "Segment",
    "Name of Ocean Vessel", "Equipment Num", "Container Size Code",
    "Carrier Name", "Site", "Route",
]

OM_COLS = [
    "PEDIDOS_MATCH[SO#]",
    "PEDIDOS_MATCH[EndCustomerName]",
    "PEDIDOS_MATCH[Customer_Name]",
    "PEDIDOS_MATCH[Material]",
    "PEDIDOS_MATCH[Order_Qty]",
    "PEDIDOS_MATCH[Open_Qty]",
    "PEDIDOS_MATCH[RPL_TRANSIT]",
    "PEDIDOS_MATCH[RPL_SO]",
    "PEDIDOS_MATCH[Sold_to]",
    "PEDIDOS_MATCH[Ctry]",
]

OM_SEMANAL_COLS = [
    "Z", "Ctry", "SO#", "Material", "Order_Qty", "Open_Qty",
    "SO Transit", "ETA_to_DC", "RPL SO", "PO Num",
    "Sold_to", "Customer_Name", "EndCustomerName", "Vta Sucesiva/DD",
]


def serial_excel(fecha: datetime | pd.Timestamp | None) -> int | None:
    if fecha is None or pd.isna(fecha):
        return None
    return int((pd.Timestamp(fecha).normalize() - pd.Timestamp("1899-12-30")).days)


def fmt_eta_texto(fecha: datetime | pd.Timestamp | None) -> str | None:
    if fecha is None or pd.isna(fecha):
        return None
    return pd.Timestamp(fecha).strftime("%m/%d/%Y")


def fmt_create_date(fecha: datetime | pd.Timestamp | None) -> str | None:
    if fecha is None or pd.isna(fecha):
        return None
    return pd.Timestamp(fecha).strftime("%Y-%m-%d")


def fila_transito(
    *,
    key: str,
    dn: str,
    item: str,
    qty: int,
    so: str,
    mtm: str,
    mot: str,
    bl: str,
    ship_date: datetime,
    eta: datetime,
    revised_eta: datetime | None = None,
    outlook_eta: datetime | None = None,
    atd: datetime | None = None,
    ata: datetime | None = None,
    create_date: datetime | None = None,
    country: str = "PE",
    storage_loc: str = "PE01",
    buque: str | None = None,
    contenedor: str | None = None,
    segment: str = "REL",
    carrier: str | None = None,
    site: str = "SINT-PLANT-CN",
    route: str = "SINT-PE01",
) -> dict:
    rev = revised_eta or eta
    outl = outlook_eta or (rev + timedelta(days=2))
    f_atd = atd or ship_date
    f_create = create_date or (ship_date - timedelta(days=3))
    es_mar = mot.upper() in {"SEA", "SEA-AIR"}
    valores = {
        "Key": key,
        "Delivery Order Number": dn,
        "Delivery Item Number": item,
        "Transit Qty": qty,
        "SO": so,
        "MTM": mtm,
        "MOT": mot,
        "Country": country,
        "Storage Loc": storage_loc,
        "BL/AWB Num": bl,
        "Ship Date": serial_excel(ship_date),
        "ETA": fmt_eta_texto(eta),
        "Revised ETA": serial_excel(rev),
        "Outlook ETA": serial_excel(outl),
        "ETD": serial_excel(ship_date),
        "ATD": serial_excel(f_atd),
        "ATA": serial_excel(ata),
        "ETA Warehouse": serial_excel(outl + timedelta(days=2)),
        "Pick Up": serial_excel(ship_date - timedelta(days=1)),
        "Create Date": fmt_create_date(f_create),
        "DC STO": f"4599{dn[-6:]}",
        "PO": f"4500{dn[-6:]}",
        "Bill": f"9099{dn[-6:]}",
        "Segment": segment,
        "Name of Ocean Vessel": buque,
        "Equipment Num": contenedor,
        "Container Size Code": "40HC" if es_mar else None,
        "Carrier Name": carrier or ("SINTLINE" if es_mar else "SINTAIR"),
        "Site": site,
        "Route": route,
    }
    return {f"Transito[{col}]": valores[col] for col in TRANSITO_COLS}


def fila_om_diario(
    so_local: str,
    cliente: str,
    mtm: str,
    rpl_transit: str = "",
    rpl_so: str = "",
    sold_to: str | None = None,
    qty: int = 20,
) -> dict:
    return {
        "PEDIDOS_MATCH[SO#]": so_local,
        "PEDIDOS_MATCH[EndCustomerName]": cliente,
        "PEDIDOS_MATCH[Customer_Name]": cliente,
        "PEDIDOS_MATCH[Material]": mtm,
        "PEDIDOS_MATCH[Order_Qty]": str(qty),
        "PEDIDOS_MATCH[Open_Qty]": str(qty),
        "PEDIDOS_MATCH[RPL_TRANSIT]": rpl_transit,
        "PEDIDOS_MATCH[RPL_SO]": rpl_so,
        "PEDIDOS_MATCH[Sold_to]": sold_to or "",
        "PEDIDOS_MATCH[Ctry]": "PE",
    }


def escribir_om_excel(ruta: Path, filas_gl: list[dict], filas_om: list[dict]) -> Path:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    df_gl = pd.DataFrame(filas_gl, columns=[f"Transito[{c}]" for c in TRANSITO_COLS])
    df_om = pd.DataFrame(filas_om, columns=OM_COLS)
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        df_gl.to_excel(w, sheet_name="Transito", index=False)
        df_om.to_excel(w, sheet_name="PEDIDOS_MATCH", index=False)
    return ruta


def construir_hoja_transit_copy_paste(filas_gl: list[dict], semana: int) -> pd.DataFrame:
    """Hoja Transit como en los correos semanales: fila 1 banner Copy & Paste / Formula, fila 2 encabezados."""
    df = pd.DataFrame(filas_gl, columns=[f"Transito[{c}]" for c in TRANSITO_COLS]).assign(Week=semana, BU="IDG")
    banner = ["Copy & Paste"] * (df.shape[1] - 2) + ["Formula", "Llenar con BU"]
    matriz = [banner, list(df.columns)] + df.astype(object).where(df.notna(), "").values.tolist()
    return pd.DataFrame(matriz)


def escribir_inbound_semanal(
    carpeta: Path,
    fecha_envio: datetime,
    semana: int,
    filas_gl: list[dict],
    filas_om_sem: list[tuple],
) -> tuple[Path, Path]:
    """Genera los 2 adjuntos semanales: Inbound mm.dd.yyyy.xlsx y Order Match mm.dd.yyyy.xlsx."""
    etiqueta = fecha_envio.strftime("%m.%d.%Y")
    ruta_inb = carpeta / f"Inbound {etiqueta}.xlsx"
    ruta_om = carpeta / f"Order Match {etiqueta}.xlsx"

    hoja_tr = construir_hoja_transit_copy_paste(filas_gl, semana)
    resumen = pd.DataFrame({"Reporte": [f"SINT Inbound Semana {semana}"], "Filas": [len(filas_gl)]})
    with pd.ExcelWriter(ruta_inb, engine="openpyxl") as w:
        resumen.to_excel(w, sheet_name="Summary", index=False)
        hoja_tr.to_excel(w, sheet_name="Transit", index=False, header=False)

    registros_om = [
        dict(zip(
            OM_SEMANAL_COLS,
            ["ZOR", "PE", so433, mtm, "20", "20", so_tr, "", rpl_so, f"PO-{so433[-4:]}",
             sold_to or "", cliente, cliente, ""],
        ))
        for so433, mtm, so_tr, rpl_so, cliente, sold_to in filas_om_sem
    ]
    df_om_sem = pd.DataFrame(registros_om, columns=OM_SEMANAL_COLS)
    with pd.ExcelWriter(ruta_om, engine="openpyxl") as w:
        pd.DataFrame({"Pais": ["PE"], "Pedidos": [len(df_om_sem)]}).to_excel(w, sheet_name="Resumen", index=False)
        df_om_sem.to_excel(w, sheet_name="Order Match", index=False)

    return ruta_inb, ruta_om


def generar_hus_lake(filas_gl: list[dict], rng: random.Random, momento: datetime) -> pd.DataFrame:
    """Construye HU sintéticas deterministas (25 columnas varchar como Lake) para cada DN del reporte."""
    df = pd.DataFrame(filas_gl)
    col_dn = "Transito[Delivery Order Number]"
    col_qty = "Transito[Transit Qty]"
    col_mot = "Transito[MOT]"
    por_dn = (
        df.dropna(subset=[col_dn])
        .groupby(col_dn, sort=True)
        .agg(unidades=(col_qty, "sum"), mot=(col_mot, "first"))
    )
    registros = []
    for idx, (dn, info) in enumerate(por_dn.iterrows(), start=1):
        es_mar = str(info["mot"]).upper() in {"SEA", "SEA-AIR"}
        n_bultos = 3 if es_mar else 2
        total_u = int(info["unidades"])
        base_u, resto_u = divmod(total_u, n_bultos)
        for b in range(n_bultos):
            u_bulto = base_u + (1 if b < resto_u else 0)
            largo = 120 if es_mar else 80
            ancho = 100 if es_mar else 60
            alto = 110 if es_mar else 70
            vol_m3 = round((largo * ancho * alto) / 1_000_000, 4)
            peso_neto = round(max(5.0, u_bulto * (2.4 if es_mar else 1.8)), 2)
            peso_bruto = round(peso_neto + (12.0 if es_mar else 4.0), 2)
            registros.append({
                "delivery_number": str(dn),
                "handling_unit_id": f"SINT-HU-{dn[-4:]}-{b + 1:02d}",
                "sscc": f"003999000{idx:05d}{b + 1:04d}",
                "batch_number": f"SINTB{idx:03d}",
                "length": str(largo),
                "width": str(ancho),
                "height": str(alto),
                "unit_of_dimension": "CMT",
                "total_volume": f"{vol_m3:.4f}",
                "volume_unit": "MTQ",
                "net_weight": f"{peso_neto:.2f}",
                "total_weight": f"{peso_bruto:.2f}",
                "chargeable_weight": f"{round(peso_bruto * 1.05, 2):.2f}",
                "volume_weight": f"{round(vol_m3 * 167.0, 2):.2f}",
                "weight_unit": "KGM",
                "tare_weight": "12.00" if es_mar else "4.00",
                "weight_unit_tare": "KGM",
                "hu_quantity": str(u_bulto),
                "delivery_deleted": "N",
                "ups_1z_number": "",
                "return_1z_number_hu": "",
                "gps_device_id": "",
                "gps_device_vendor_id": "",
                "gps_shipment_id": "",
                "loc_updated_time": momento.strftime("%Y-%m-%d %H:%M:%S"),
            })
        # Para la primera DN agregamos 1 HU eliminada ('X') para ejercitar job_in_transit.eliminadas()
        if idx == 1:
            eliminada = dict(registros[-1])
            eliminada["handling_unit_id"] = f"SINT-HU-{dn[-4:]}-99"
            eliminada["delivery_deleted"] = "X"
            registros.append(eliminada)
    return pd.DataFrame(registros, columns=HU_COLS_25)


def generar_mapa_topojson_minimo(ruta: Path) -> Path:
    """TopoJSON mínimo válido sin red (incluye Perú id='604' alrededor del Callao y costa vecina)."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    topo = {
        "type": "Topology",
        "transform": {"scale": [1.0, 1.0], "translate": [0.0, 0.0]},
        "objects": {
            "countries": {
                "type": "GeometryCollection",
                "geometries": [
                    {"type": "Polygon", "id": "604", "arcs": [[0]]},
                    {"type": "Polygon", "id": "218", "arcs": [[1]]},
                    {"type": "Polygon", "id": "152", "arcs": [[2]]},
                ],
            }
        },
        "arcs": [
            [[-81.3, -3.5], [12.0, 0.0], [0.0, -14.8], [-9.0, 0.0], [-3.0, 14.8]],
            [[-81.0, 1.4], [6.0, 0.0], [0.0, -4.9], [-6.0, 0.0], [0.0, 4.9]],
            [[-70.5, -18.3], [4.0, 0.0], [0.0, -18.0], [-4.0, 0.0], [0.0, 18.0]],
        ],
    }
    ruta.write_text(json.dumps(topo, separators=(",", ":")), encoding="utf-8")
    return ruta


def html_programacion_fwd(filas_bl_awb: list[tuple[str, str, int, int]]) -> str:
    """Tabla HTML de 'PROGRAMACION DE ENTREGA CLIENTE' como la que parsea extraer_tabla_html."""
    trs = "".join(
        f"<tr><td>{i}</td><td>CLIENTE PERU</td><td>{bl}</td><td>{awb}</td>"
        f"<td>{bultos}</td><td>{peso}</td><td>TERMINAL CALLAO</td><td>DC Lima</td></tr>"
        for i, (bl, awb, bultos, peso) in enumerate(filas_bl_awb, start=1)
    )
    return (
        "<html><body><p>Estimados, adjuntamos cuadro de despacho.</p>"
        "<table border='1'>"
        "<tr><th>ORDEN</th><th>CLIENTE</th><th>BL</th><th>AWB</th>"
        "<th>BULTOS</th><th>PESO</th><th>TERMINAL</th><th>DESTINO</th></tr>"
        f"{trs}</table><p>Saludos cordiales,</p></body></html>"
    )


def construir_escenarios_y_datos(fecha_hoy_str: str = FECHA_HOY_DEFAULT, semilla: int = SEMILLA_DEFAULT) -> EstadoFakes:
    """Limpia y puebla sandbox/ de forma 100% determinista y retorna el EstadoFakes listo para correr."""
    rng = random.Random(semilla)
    np.random.seed(semilla % (2**32 - 1))

    hoy = datetime.strptime(fecha_hoy_str, "%Y-%m-%d")
    ayer = hoy - timedelta(days=1)
    d = lambda dias, h=0, m=0: hoy + timedelta(days=dias, hours=h, minutes=m)

    if SANDBOX.exists():
        shutil.rmtree(SANDBOX)
    for carpeta in (DIR_INBOUND, DIR_TRACKING, DIR_M4, DIR_INSUMOS):
        carpeta.mkdir(parents=True, exist_ok=True)

    # ── 1. Líneas de Transito para Día 1 (ayer) y Día 2 (hoy) ──────────────
    # ESC-01: Aéreo en tránsito internacional (2 ítems en la misma DN -> ejercita volumen repartido)
    k01a = fila_transito(
        key="SINT-K01A", dn="8900001001", item="10", qty=15, so="4229900101", mtm="SINT-MTM-01",
        mot="AIR", bl="SINTA-100001", ship_date=d(-2), eta=d(3), outlook_eta=d(4),
    )
    k01b = fila_transito(
        key="SINT-K01B", dn="8900001001", item="20", qty=10, so="4229900101", mtm="SINT-MTM-01B",
        mot="AIR", bl="SINTA-100001", ship_date=d(-2), eta=d(3), outlook_eta=d(4),
    )

    # ESC-02A: Aéreo en Día 1 que SALE de Transito en Día 2; tiene aviso con PDF texto (fuera SLA + 1 bulto malo)
    k02 = fila_transito(
        key="SINT-K02", dn="8900001002", item="10", qty=30, so="4229900102", mtm="SINT-MTM-02",
        mot="AIR", bl="SINTA-100002", ship_date=d(-11), eta=d(-8), outlook_eta=d(-7),
    )

    # ESC-02B: Aéreo en Día 1 que SALE de Transito en Día 2; tiene aviso con PDF escaneado sin texto
    k03 = fila_transito(
        key="SINT-K03", dn="8900001003", item="10", qty=15, so="4229900103", mtm="SINT-MTM-03",
        mot="AIR", bl="SINTA-100003", ship_date=d(-6), eta=d(-2), outlook_eta=d(-1),
    )

    # ESC-03: Aéreo con aviso + salida FWD programada hoy (En tránsito a DC / En ruta al DC)
    k04 = fila_transito(
        key="SINT-K04", dn="8900001004", item="10", qty=40, so="4229900104", mtm="SINT-MTM-04",
        mot="AIR", bl="SINTA-100004", ship_date=d(-7), eta=d(-3), outlook_eta=d(-2),
    )

    # ESC-04: Aéreo entregado al DC (con recepción 3PL + 433 con DN en M4 con GI)
    k05 = fila_transito(
        key="SINT-K05", dn="8900001005", item="10", qty=18, so="4229900105", mtm="SINT-MTM-05",
        mot="AIR", bl="SINTA-100005", ship_date=d(-10), eta=d(-6), outlook_eta=d(-5),
    )

    # ESC-05: Marítimo en tránsito con buque cercano ("SINT AURORA") + ítem rama Panamá (PA/PE01)
    k06 = fila_transito(
        key="SINT-K06", dn="8900002001", item="10", qty=150, so="4229900201", mtm="SINT-MTM-06",
        mot="SEA", bl="SINTM-200001", ship_date=d(-25), eta=d(2), revised_eta=d(2), outlook_eta=d(4),
        buque="SINT AURORA", contenedor="SINTU200001",
    )
    k06_pa = fila_transito(
        key="SINT-K06PA", dn="8900002091", item="10", qty=20, so="4229900201", mtm="SINT-MTM-06",
        mot="SEA", bl="SINTM-200001", ship_date=d(-25), eta=d(2), revised_eta=d(2), outlook_eta=d(4),
        country="PA", storage_loc="PE01", buque="SINT AURORA", contenedor="SINTU200001",
    )

    # ESC-06: Marítimo arribado al terminal que YA NO está en Transito de ayer ni hoy, pero SÍ en Inbound semanal
    k06_perdido_sem = fila_transito(
        key="SINT-K06M-SEM", dn="8900002002", item="10", qty=200, so="4229900202", mtm="SINT-MTM-07",
        mot="SEA", bl="SINTM-200002", ship_date=d(-30), eta=d(-2), revised_eta=d(-2), outlook_eta=d(-1),
        buque="SINT PACIFICO", contenedor="SINTU200002",
    )

    # ESC-07: Marítimo arribado con salida a DC programada hoy
    k07 = fila_transito(
        key="SINT-K07", dn="8900002003", item="10", qty=180, so="4229900203", mtm="SINT-MTM-08",
        mot="SEA", bl="SINTM-200003", ship_date=d(-28), eta=d(-4), revised_eta=d(-3), outlook_eta=d(-2),
        ata=d(-3), buque="SINT CALLAO EXPRESS", contenedor="SINTU200003",
    )

    # ESC-08: Marítimo cuya ETA se mueve >5 días entre semanas / entre Día 1 (d+1) y Día 2 (d+10: +9 días)
    k08_sem1 = fila_transito(
        key="SINT-K08", dn="8900002004", item="10", qty=120, so="4229900204", mtm="SINT-MTM-09",
        mot="SEA", bl="SINTM-200004", ship_date=d(-24), eta=d(-4), revised_eta=d(-4), outlook_eta=d(-2),
        buque="SINT HORIZONTE", contenedor="SINTU200004",
    )
    k08_sem2 = fila_transito(
        key="SINT-K08", dn="8900002004", item="10", qty=120, so="4229900204", mtm="SINT-MTM-09",
        mot="SEA", bl="SINTM-200004", ship_date=d(-24), eta=d(-4), revised_eta=d(-1), outlook_eta=d(1),
        buque="SINT HORIZONTE", contenedor="SINTU200004",
    )
    k08_dia1 = fila_transito(
        key="SINT-K08", dn="8900002004", item="10", qty=120, so="4229900204", mtm="SINT-MTM-09",
        mot="SEA", bl="SINTM-200004", ship_date=d(-24), eta=d(1), revised_eta=d(1), outlook_eta=d(3),
        buque="SINT HORIZONTE", contenedor="SINTU200004",
    )
    k08_dia2 = fila_transito(
        key="SINT-K08", dn="8900002004", item="10", qty=120, so="4229900204", mtm="SINT-MTM-09",
        mot="SEA", bl="SINTM-200004", ship_date=d(-24), eta=d(10), revised_eta=d(10), outlook_eta=d(12),
        buque="SINT HORIZONTE", contenedor="SINTU200004",
    )

    # ESC-09: Línea de stock sin 433 (en bi) + ESC-09B: aéreo de stock rescatado de Inbound semanal
    k09 = fila_transito(
        key="SINT-K09", dn="8900003001", item="10", qty=50, so="4229900301", mtm="SINT-MTM-10",
        mot="AIR", bl="SINTA-100006", ship_date=d(-5), eta=d(-1), outlook_eta=d(0),
    )
    k09_sem = fila_transito(
        key="SINT-K09-SEM", dn="8900003002", item="10", qty=22, so="4229900302", mtm="SINT-MTM-10",
        mot="AIR", bl="SINTA-100009", ship_date=d(-14), eta=d(-3), outlook_eta=d(-2),
    )

    # ESC-10: 422 que abastece DOS pedidos 433 distintos
    k10 = fila_transito(
        key="SINT-K10", dn="8900004001", item="10", qty=60, so="4229900401", mtm="SINT-MTM-11",
        mot="AIR", bl="SINTA-100007", ship_date=d(-5), eta=d(-1), outlook_eta=d(0),
    )

    # ESC-11: Embarque consolidado con DOS MOT distintos (AIR + SEA en el mismo HAWB/BL)
    k11_air = fila_transito(
        key="SINT-K11A", dn="8900005001", item="10", qty=20, so="4229900501", mtm="SINT-MTM-12",
        mot="AIR", bl="SINTA-100008", ship_date=d(-15), eta=d(-12), outlook_eta=d(-11),
    )
    k11_sea = fila_transito(
        key="SINT-K11B", dn="8900005002", item="10", qty=30, so="4229900502", mtm="SINT-MTM-13",
        mot="SEA", bl="SINTA-100008", ship_date=d(-32), eta=d(-12), revised_eta=d(-12), outlook_eta=d(-11),
        ata=d(-12), buque="SINT CONDOR", contenedor="SINTU200008",
    )

    # ESC-12: Marítimo entregado al DC (salida FWD > 2 d.h.) con una 433 SIN DN en M4 y otra 433 CON DN en M4 sin GI
    k12a = fila_transito(
        key="SINT-K12A", dn="8900006001", item="10", qty=70, so="4229900601", mtm="SINT-MTM-14",
        mot="SEA", bl="SINTM-200005", ship_date=d(-35), eta=d(-10), revised_eta=d(-9), outlook_eta=d(-8),
        ata=d(-9), buque="SINT CONDOR", contenedor="SINTU200005",
    )
    k12b = fila_transito(
        key="SINT-K12B", dn="8900006002", item="10", qty=45, so="4229900602", mtm="SINT-MTM-15",
        mot="SEA", bl="SINTM-200005", ship_date=d(-35), eta=d(-10), revised_eta=d(-9), outlook_eta=d(-8),
        ata=d(-9), buque="SINT CONDOR", contenedor="SINTU200005",
    )

    # Línea PA fuera de alcance (PA01 -> filtrada por job_in_transit)
    k_pa_fuera = fila_transito(
        key="SINT-K99PA", dn="8900009999", item="10", qty=12, so="4229900999", mtm="SINT-MTM-99",
        mot="SEA", bl="SINTM-209999", ship_date=d(-10), eta=d(8), country="PA", storage_loc="PA01",
        buque="SINT LEJANO", contenedor="SINTU209999",
    )

    gl_dia1 = [
        k01a, k01b, k02, k03, k04, k05,
        k06, k06_pa, k07, k08_dia1,
        k09, k10, k11_air, k11_sea, k12a, k12b, k_pa_fuera,
    ]
    # En Día 2 salen de Transito los aéreos k02 (SINTA-100002) y k03 (SINTA-100003), y k08 mueve su ETA +9 días
    gl_dia2 = [
        k01a, k01b, k04, k05,
        k06, k06_pa, k07, k08_dia2,
        k09, k10, k11_air, k11_sea, k12a, k12b, k_pa_fuera,
    ]

    om_comunes = [
        fila_om_diario("4339900101", "SINT BANCO ANDINO SAC", "SINT-MTM-01", rpl_transit="4229900101", sold_to="SINT-ST-101", qty=15),
        fila_om_diario("4339900101", "SINT BANCO ANDINO SAC", "SINT-MTM-01B", rpl_transit="4229900101", sold_to="SINT-ST-101", qty=10),
        fila_om_diario("4339900104", "SINT RETAIL PACIFICO SAC", "SINT-MTM-04", rpl_transit="4229900104", sold_to="SINT-ST-104", qty=40),
        fila_om_diario("4339900105", "SINT HOSPITAL CENTRAL", "SINT-MTM-05", rpl_transit="4229900105", sold_to="SINT-ST-105", qty=18),
        fila_om_diario("4339900201", "SINT CORPORACION ANDINA SAC", "SINT-MTM-06", rpl_transit="4229900201", sold_to="SINT-ST-201", qty=170),
        fila_om_diario("4339900203", "SINT CONSTRUCTORA LIMA SA", "SINT-MTM-08", rpl_transit="4229900203", sold_to="SINT-ST-203", qty=180),
        fila_om_diario("4339900204", "SINT LOGISTICA CALLAO SAC", "SINT-MTM-09", rpl_transit="4229900204", sold_to="SINT-ST-204", qty=120),
        # ESC-10: una misma 422 (4229900401) con dos 433 distintos (uno en RPL_TRANSIT y otro en RPL_SO)
        fila_om_diario("4339900401", "SINT UNIVERSIDAD CENTRAL", "SINT-MTM-11", rpl_transit="4229900401", sold_to="SINT-ST-401", qty=35),
        fila_om_diario("4339900402", "SINT COLEGIO TECNICO SAC", "SINT-MTM-11", rpl_so="4229900401", sold_to="SINT-ST-402", qty=25),
        # ESC-11: dos MOT
        fila_om_diario("4339900501", "SINT AGRO EXPORT SAC", "SINT-MTM-12", rpl_transit="4229900501", sold_to="SINT-ST-501", qty=20),
        fila_om_diario("4339900502", "SINT TEXTIL LIMA SAC", "SINT-MTM-13", rpl_transit="4229900502", sold_to="SINT-ST-502", qty=30),
        # ESC-12: 433 sin DN y 433 con DN en M4
        fila_om_diario("4339900601", "SINT ENERGIA PERU SA", "SINT-MTM-14", rpl_transit="4229900601", sold_to="SINT-ST-601", qty=70),
        fila_om_diario("4339900602", "SINT TELECOM ANDES SAC", "SINT-MTM-15", rpl_transit="4229900602", sold_to="SINT-ST-602", qty=45),
    ]
    # En Día 1 están los 433 de SINTA-100002 y SINTA-100003; en Día 2 su match de tránsito ya salió del OM de hoy
    om_dia1 = om_comunes + [
        fila_om_diario("4339900102", "SINT CLINICA NORTE SAC", "SINT-MTM-02", rpl_transit="4229900102", sold_to="SINT-ST-102", qty=30),
        fila_om_diario("4339900103", "SINT FARMACIAS SUR SAC", "SINT-MTM-03", rpl_transit="4229900103", sold_to="SINT-ST-103", qty=15),
    ]
    om_dia2 = list(om_comunes)

    ruta_om_d1 = escribir_om_excel(DIR_INSUMOS / "OM_dia1.xlsx", gl_dia1, om_dia1)
    ruta_om_d2 = escribir_om_excel(DIR_INSUMOS / "OM_dia2.xlsx", gl_dia2, om_dia2)
    shutil.copy2(ruta_om_d1, DIR_INBOUND / "OM.xlsx")

    # ── 2. HUs sintéticas de Lake por día ─────────────────────────────────────
    hus_d1 = generar_hus_lake(gl_dia1, rng, ayer.replace(hour=8, minute=0))
    hus_d2 = generar_hus_lake(gl_dia2, rng, hoy.replace(hour=8, minute=0))

    # ── 3. M4 parquets mínimos (deliverys_&_returns) ──────────────────────────
    m4_con_gi = pd.DataFrame([
        {
            "site_country": "PERU",
            "so": "4339900105",
            "dn": "8599000105",
            "dn_create_date": pd.Timestamp(d(-1)),
            "dn_gi_date": pd.Timestamp(d(0)),
        },
        {
            "site_country": "CHILE",
            "so": "4339999999",
            "dn": "8599999999",
            "dn_create_date": pd.Timestamp(d(-3)),
            "dn_gi_date": pd.Timestamp(d(-2)),
        },
    ])
    m4_sin_gi = pd.DataFrame([
        {
            "site_country": "PERU",
            "so": "4339900602",
            "dn": "8599000602",
            "dn_create_date": pd.Timestamp(d(0)),
            "dn_gi_date": pd.NaT,
        },
    ])
    m4_con_gi.to_parquet(DIR_M4 / "m4_latam.parquet", index=False)
    m4_sin_gi.to_parquet(DIR_M4 / "m4_latam_sin_gi.parquet", index=False)

    # ── 4. Mapa mundial mínimo falso (sin red) ────────────────────────────────
    generar_mapa_topojson_minimo(DIR_TRACKING / "mapa_mundial_50m.json")

    # ── 5. Tres correos semanales "RE: Inbound & Order Match" en Enviados ─────
    dir_sem = DIR_INSUMOS / "semanales"
    dir_sem.mkdir(parents=True, exist_ok=True)

    f_sem1 = d(-19, 17, 0)
    inb_s1, om_s1 = escribir_inbound_semanal(
        dir_sem, f_sem1, 45,
        [k08_sem1, k12a, k12b],
        [
            ("4339900204", "SINT-MTM-09", "4229900204", "", "SINT LOGISTICA CALLAO SAC", "SINT-ST-204"),
            ("4339900601", "SINT-MTM-14", "4229900601", "", "SINT ENERGIA PERU SA", "SINT-ST-601"),
            ("4339900602", "SINT-MTM-15", "4229900602", "", "SINT TELECOM ANDES SAC", "SINT-ST-602"),
        ],
    )

    f_sem2 = d(-12, 17, 0)
    inb_s2, om_s2 = escribir_inbound_semanal(
        dir_sem, f_sem2, 46,
        [k05, k06, k07, k08_sem2, k09_sem, k11_air, k11_sea],
        [
            ("4339900105", "SINT-MTM-05", "4229900105", "", "SINT HOSPITAL CENTRAL", "SINT-ST-105"),
            ("4339900201", "SINT-MTM-06", "4229900201", "", "SINT CORPORACION ANDINA SAC", "SINT-ST-201"),
            ("4339900203", "SINT-MTM-08", "4229900203", "", "SINT CONSTRUCTORA LIMA SA", "SINT-ST-203"),
            ("4339900204", "SINT-MTM-09", "4229900204", "", "SINT LOGISTICA CALLAO SAC", "SINT-ST-204"),
        ],
    )

    f_sem3 = d(-5, 17, 0)
    inb_s3, om_s3 = escribir_inbound_semanal(
        dir_sem, f_sem3, 47,
        [k01a, k01b, k02, k03, k04, k06, k06_perdido_sem, k07, k08_dia1, k10],
        [
            ("4339900101", "SINT-MTM-01", "4229900101", "", "SINT BANCO ANDINO SAC", "SINT-ST-101"),
            ("4339900102", "SINT-MTM-02", "4229900102", "", "SINT CLINICA NORTE SAC", "SINT-ST-102"),
            ("4339900103", "SINT-MTM-03", "4229900103", "", "SINT FARMACIAS SUR SAC", "SINT-ST-103"),
            ("4339900202", "SINT-MTM-07", "4229900202", "", "SINT MINERA DEL SUR SA", "SINT-ST-202"),
            ("4339900204", "SINT-MTM-09", "4229900204", "", "SINT LOGISTICA CALLAO SAC", "SINT-ST-204"),
        ],
    )

    enviados = [
        CorreoFalso("<sint-sem1@sint.local>", "RE: Inbound & Order Match", "Envío semanal 1", f_sem1,
                    adjuntos=[AdjuntoMail(inb_s1), AdjuntoMail(om_s1)]),
        CorreoFalso("<sint-sem2@sint.local>", "RE: Inbound & Order Match", "Envío semanal 2", f_sem2,
                    adjuntos=[AdjuntoMail(inb_s2), AdjuntoMail(om_s2)]),
        CorreoFalso("<sint-sem3@sint.local>", "RE: Inbound & Order Match", "Envío semanal 3", f_sem3,
                    adjuntos=[AdjuntoMail(inb_s3), AdjuntoMail(om_s3)]),
    ]

    # ── 6. PDFs y correos entrantes de Outlook (Inbox + Shipments) ────────────
    dir_pdfs = DIR_INSUMOS / "pdfs"
    pdf_k02 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100002.pdf",
        [f"TARJA TERMINAL AEREO\nFecha de llegada: {d(-8):%d/%m/%Y}\nBultos 4 0 3 1 0"],
    )
    pdf_k03_escaneado = escribir_pdf_sintetico(
        dir_pdfs / "aviso_escaneado_SINTA-100003.pdf",
        [""],  # PDF escaneado sin texto extraíble
    )
    pdf_k04 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100004.pdf",
        [f"Fecha de llegada: {d(-3):%d/%m/%Y}\nBultos 5 0 5 0 0"],
    )
    pdf_k05 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100005.pdf",
        [f"Fecha de llegada: {d(-6):%d/%m/%Y}\nBultos 2 0 2 0 0"],
    )
    pdf_k09 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100006.pdf",
        [f"Fecha de llegada: {d(-1):%d/%m/%Y}\nBultos 2 0 2 0 0"],
    )
    pdf_k09_sem = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100009.pdf",
        [f"Fecha de llegada: {d(-2):%d/%m/%Y}\nBultos 2 0 2 0 0"],
    )
    pdf_k10 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100007.pdf",
        [f"Fecha de llegada: {d(-1):%d/%m/%Y}\nBultos 2 0 2 0 0"],
    )
    pdf_k11 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_SINTA-100008.pdf",
        [f"Fecha de llegada: {d(-12):%d/%m/%Y}\nBultos 4 0 4 0 0"],
    )
    pdf_mar_k06 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_mar_SINTM-200001.pdf",
        [f"Pto. Descarga CALLAO   ETA/ATA {d(2):%d.%m.%Y}\nETD/ATD {d(-25):%d.%m.%Y}"],
    )
    pdf_mar_k06_perdido = escribir_pdf_sintetico(
        dir_pdfs / "aviso_mar_SINTM-200002.pdf",
        [f"Pto. Descarga CALLAO   ETA/ATA {d(-2):%d.%m.%Y}\nETD/ATD {d(-30):%d.%m.%Y}"],
    )
    pdf_mar_k07 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_mar_SINTM-200003.pdf",
        [f"Pto. Descarga CALLAO   ETA/ATA {d(-3):%d.%m.%Y}\nETD/ATD {d(-28):%d.%m.%Y}"],
    )
    pdf_mar_k12 = escribir_pdf_sintetico(
        dir_pdfs / "aviso_mar_SINTM-200005.pdf",
        [f"Pto. Descarga CALLAO   ETA/ATA {d(-9):%d.%m.%Y}\nETD/ATD {d(-35):%d.%m.%Y}"],
    )

    cita_fwd = lambda txt: (
        f"Estimados,\n\nFrom: FWD Origin\nSent: Monday, November 2, 2026 09:00 AM\n"
        f"To: cliente\nSubject: Shipment\n{txt}"
    )

    inbox = [
        # Avisos aéreos disponibles ya en Día 1 (<= ayer 08:30)
        CorreoFalso("<sint-arr-100003@sint.local>",
                    "[External] [#TN#9901 #LIM] AVISO DE LLEGADA : SINTA-100003 || CLIENTE",
                    "Adjuntamos aviso escaneado.", d(-2, 14, 0), adjuntos=[AdjuntoMail(pdf_k03_escaneado)]),
        CorreoFalso("<sint-arr-100004@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100004",
                    "Adjuntamos tarja.", d(-3, 10, 0), adjuntos=[AdjuntoMail(pdf_k04)]),
        CorreoFalso("<sint-arr-100005@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100005",
                    "Adjuntamos tarja.", d(-6, 9, 30), adjuntos=[AdjuntoMail(pdf_k05)]),
        CorreoFalso("<sint-arr-100009@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100009",
                    "Adjuntamos tarja.", d(-2, 11, 0), adjuntos=[AdjuntoMail(pdf_k09_sem)]),
        CorreoFalso("<sint-arr-100008@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100008",
                    "Adjuntamos tarja.", d(-12, 10, 0), adjuntos=[AdjuntoMail(pdf_k11)]),

        #Avisos aéreos que llegan después del corte de Día 1 (ayer > 08:30) -> entran en Día 2
        CorreoFalso("<sint-arr-100002@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100002",
                    "Adjuntamos tarja con observación.", d(-1, 11, 30), adjuntos=[AdjuntoMail(pdf_k02)]),
        CorreoFalso("<sint-arr-100006@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100006",
                    "Adjuntamos tarja.", d(-1, 12, 0), adjuntos=[AdjuntoMail(pdf_k09)]),
        CorreoFalso("<sint-arr-100007@sint.local>",
                    "AVISO DE LLEGADA HAWB SINTA-100007",
                    "Adjuntamos tarja.", d(-1, 12, 30), adjuntos=[AdjuntoMail(pdf_k10)]),

        # Programación FWD antigua (>2 d.h.) -> Entregado al DC para SINTA-100005 y SINTM-200005
        CorreoFalso(
            "<sint-prog-antigua@sint.local>",
            f"PROGRAMACION DE ENTREGA CLIENTE {d(-7):%d/%m}",
            f"Estimados, enviamos programación para el día {d(-7):%d/%m} a las 08:30 Hrs.",
            d(-7, 16, 0),
            html_body=html_programacion_fwd([
                ("SINTM-200005", "", 6, 980),
                ("", "SINTA-100005", 2, 75),
            ]),
        ),
        # Programación FWD vigente para HOY (llega en la tarde de ayer -> entra en Día 2 para SINTA-100004 y SINTM-200003)
        CorreoFalso(
            "<sint-prog-hoy@sint.local>",
            f"PROGRAMACION DE ENTREGA CLIENTE {d(0):%d/%m}",
            f"Estimados, enviamos programación para el día {d(0):%d/%m} a las 10:30 Hrs.",
            d(-1, 16, 30),
            html_body=html_programacion_fwd([
                ("SINTM-200003", "", 3, 520),
                ("", "SINTA-100004", 2, 110),
            ]),
        ),

        # Recepción 3PL para SINTA-100005
        CorreoFalso(
            "<sint-op3pl-100005@sint.local>",
            f"IMPORTACION CLIENTE - Recepción {d(-2):%d/%m/%Y}",
            "Confirmamos ingreso en almacén DC Lima de guía SINTA-100005.",
            d(-2, 17, 30),
        ),

        # Correos marítimos FWD (prealerta, notificación, aviso con buque/POL/ETA/ATA/ATD)
        CorreoFalso(
            "<sint-kn-pre-200001@sint.local>",
            "Cliente Shipment PREALERT SINTM-200001",
            f"ATD: {d(-25):%Y-%m-%d}\nETA: {d(3):%Y-%m-%d}\nP. of Discharge: CALLAO",
            d(-20, 9, 0),
        ),
        CorreoFalso(
            "<sint-kn-av-200001@sint.local>",
            "AVISO DE LLEGADA // HBL SINTM-200001 // POL: CNSHA // BUQUE: SINT AURORA",
            cita_fwd(f"ATD: {d(-25):%Y-%m-%d}\nETA: {d(3):%Y-%m-%d}"),
            d(-4, 15, 0),
            adjuntos=[AdjuntoMail(pdf_mar_k06)],
        ),
        CorreoFalso(
            "<sint-kn-not-200001@sint.local>",
            "Cliente Shipment / SINTM-200001",
            f"New ETA: {d(2):%Y-%m-%d}\nP. of Discharge: CALLAO",
            d(-3, 10, 30),
        ),
        # ESC-06: Aviso marítimo de SINTM-200002 (ya salió de Transito, ATA hace 2 días, buque SINT PACIFICO)
        CorreoFalso(
            "<sint-kn-av-200002@sint.local>",
            "AVISO DE LLEGADA // HBL SINTM-200002 // POL: CNSHK // BUQUE: SINT PACIFICO",
            cita_fwd(f"ATD: {d(-30):%Y-%m-%d}"),
            d(-2, 10, 0),
            adjuntos=[AdjuntoMail(pdf_mar_k06_perdido)],
        ),
        # ESC-07: Aviso marítimo de SINTM-200003 (ATA hace 3 días, buque SINT CALLAO EXPRESS)
        CorreoFalso(
            "<sint-kn-av-200003@sint.local>",
            "AVISO DE LLEGADA // HBL SINTM-200003 // POL: CNSHA // BUQUE: SINT CALLAO EXPRESS",
            cita_fwd(f"ATD: {d(-28):%Y-%m-%d}"),
            d(-3, 14, 0),
            adjuntos=[AdjuntoMail(pdf_mar_k07)],
        ),
        # ESC-08: Prealerta + notificación Día 1 + respuesta Día 2 que mueve ETA +9 días
        CorreoFalso(
            "<sint-kn-pre-200004@sint.local>",
            "Cliente Shipment PREALERT SINTM-200004",
            f"ATD: {d(-24):%Y-%m-%d}\nETA: {d(-4):%Y-%m-%d}\nP. of Discharge: CALLAO",
            d(-22, 8, 0),
        ),
        CorreoFalso(
            "<sint-kn-not-200004-d1@sint.local>",
            "Cliente Shipment / SINTM-200004",
            f"New ETA: {d(1):%Y-%m-%d}\nP. of Discharge: CALLAO",
            d(-6, 9, 0),
        ),
        CorreoFalso(
            "<sint-kn-re-200004-d2@sint.local>",
            "RE: Cliente Shipment / SINTM-200004",
            f"Estimados, naviera reprograma New ETA: {d(10):%Y-%m-%d}\n\n"
            + cita_fwd(f"ATD: {d(-24):%Y-%m-%d}\nETA: {d(1):%Y-%m-%d}"),
            d(0, 7, 0),
        ),
        # ESC-12: Aviso marítimo de SINTM-200005
        CorreoFalso(
            "<sint-kn-av-200005@sint.local>",
            "AVISO DE LLEGADA // HBL SINTM-200005 // POL: CNSHA // BUQUE: SINT CONDOR",
            cita_fwd(f"ATD: {d(-35):%Y-%m-%d}"),
            d(-9, 13, 0),
            adjuntos=[AdjuntoMail(pdf_mar_k12)],
        ),
    ]
    # Copia de un correo en Shipments para ejercitar deduplicación por Internet Message-ID
    shipments = [inbox[0]]

    # ── 7. Posiciones AIS falsas por día (sin llamadas reales) ────────────────
    iso_utc = lambda dt_ref: dt_ref.strftime("%Y-%m-%dT13:00:00Z")
    barcos_ais = {
        1: {
            "SINT AURORA": {
                "nombre": "SINT AURORA", "imo": "9990001", "mmsi": "777000001",
                "lat": -9.50, "lon": -79.20, "destino": "PECLL",
                "eta_iso": iso_utc(d(2)), "actualizado_iso": iso_utc(ayer),
                "escalas_pasadas": [{"portName": "Guayaquil", "portUnlocode": "ECGYE",
                                     "arrivalTime": iso_utc(d(-4)), "departureTime": iso_utc(d(-3))}],
                "escalas_futuras": [{"port_name": "Callao", "port_unlocode": "PECLL",
                                     "eta": d(2).strftime("%Y-%m-%d"), "eta_time": "08:00:00", "is_next_port": True}],
            },
            "SINT HORIZONTE": {
                "nombre": "SINT HORIZONTE", "imo": "9990004", "mmsi": "777000004",
                "lat": 8.80, "lon": -79.55, "destino": "PABLB",
                "eta_iso": iso_utc(d(1)), "actualizado_iso": iso_utc(ayer),
                "escalas_pasadas": [{"portName": "Manzanillo", "portUnlocode": "MXZLO",
                                     "arrivalTime": iso_utc(d(-7)), "departureTime": iso_utc(d(-6))}],
                "escalas_futuras": [
                    {"port_name": "Balboa", "port_unlocode": "PABLB",
                     "eta": d(1).strftime("%Y-%m-%d"), "eta_time": "10:00:00", "is_next_port": True},
                    {"port_name": "Callao", "port_unlocode": "PECLL",
                     "eta": d(10).strftime("%Y-%m-%d"), "eta_time": "08:00:00", "is_next_port": False},
                ],
            },
        },
        2: {
            "SINT AURORA": {
                "nombre": "SINT AURORA", "imo": "9990001", "mmsi": "777000001",
                "lat": -10.80, "lon": -78.10, "destino": "PECLL",
                "eta_iso": iso_utc(d(2)), "actualizado_iso": iso_utc(hoy),
                "escalas_pasadas": [{"portName": "Guayaquil", "portUnlocode": "ECGYE",
                                     "arrivalTime": iso_utc(d(-4)), "departureTime": iso_utc(d(-3))}],
                "escalas_futuras": [{"port_name": "Callao", "port_unlocode": "PECLL",
                                     "eta": d(2).strftime("%Y-%m-%d"), "eta_time": "08:00:00", "is_next_port": True}],
            },
        },
    }

    # Persistir manifiesto de escenarios para verificación determinista en resumen_base.py
    manifiesto = {
        "fecha_hoy": fecha_hoy_str,
        "fecha_ayer": ayer.strftime("%Y-%m-%d"),
        "semilla": semilla,
        "escenarios": [
            {"id": "ESC-01", "titulo": "Aéreo en tránsito internacional", "hawb_bl": "SINTA-100001", "so": "4229900101",
             "so_local": ["4339900101"], "esperado_correo_hoy": False, "esperado_radar_hoy": False,
             "deberia_v2": "En seguimiento como En tránsito internacional; no entra a correo de arribados ni pestaña aérea de terminal."},
            {"id": "ESC-02A", "titulo": "Aéreo arribado fuera de Transito (PDF texto, fuera SLA + bulto malo)", "hawb_bl": "SINTA-100002", "so": "4229900102",
             "so_local": ["4339900102"], "esperado_correo_hoy": True, "esperado_radar_hoy": True,
             "deberia_v2": "Aparecer Arribado en correo (con Requiere acción por SLA y bulto malo, rescatando 433 + Sold-to del historial) y en pestaña Aéreo del radar."},
            {"id": "ESC-02B", "titulo": "Aéreo arribado fuera de Transito (PDF escaneado sin texto)", "hawb_bl": "SINTA-100003", "so": "4229900103",
             "so_local": ["4339900103"], "esperado_correo_hoy": True, "esperado_radar_hoy": True,
             "deberia_v2": "Aparecer Arribado con fecha del correo como respaldo, sin re-descargar el mismo PDF escaneado todos los días."},
            {"id": "ESC-03", "titulo": "Aéreo con salida FWD programada a DC", "hawb_bl": "SINTA-100004", "so": "4229900104",
             "so_local": ["4339900104"], "esperado_correo_hoy": True, "esperado_radar_hoy": False,
             "deberia_v2": "Aparecer En ruta al DC en el correo; en el radar actual solo suma en_ruta=1 al pie (en v2 entra a pestaña Salida a DC)."},
            {"id": "ESC-04", "titulo": "Aéreo entregado (recepción 3PL)", "hawb_bl": "SINTA-100005", "so": "4229900105",
             "so_local": ["4339900105"], "esperado_correo_hoy": False, "esperado_radar_hoy": False,
             "deberia_v2": "Hito Entregado al DC; hoy no sale en correo ni radar (en v2 aparece en pestaña Salida a DC como Ingresado en SAP por DN 8599000105 en M4)."},
            {"id": "ESC-05", "titulo": "Marítimo en tránsito con buque cercano (SINT AURORA)", "hawb_bl": "SINTM-200001", "so": "4229900201",
             "so_local": ["4339900201"], "esperado_correo_hoy": False, "esperado_radar_hoy": True,
             "deberia_v2": "No entra al correo (aún en tránsito marítimo); sí aparece en pestaña Marítimo del radar bajo buque SINT AURORA con posición AIS cercana."},
            {"id": "ESC-06", "titulo": "Marítimo arribado al terminal fuera de Transito pero en Inbound semanal", "hawb_bl": "SINTM-200002", "so": "4229900202",
             "so_local": ["4339900202"], "esperado_correo_hoy": False, "esperado_radar_hoy": False,
             "deberia_v2": "BRECHA CRÍTICA HOY: el backfill lo rescata pero pierde name_of_ocean_vessel (SINT PACIFICO); el correo lo excluye (~_mar sin salida DC) y el radar lo pierde por leer solo seguimiento. En v2 debe aparecer arribado en terminal con su buque y su 433."},
            {"id": "ESC-07", "titulo": "Marítimo con salida a DC programada", "hawb_bl": "SINTM-200003", "so": "4229900203",
             "so_local": ["4339900203"], "esperado_correo_hoy": True, "esperado_radar_hoy": False,
             "deberia_v2": "Aparece En ruta al DC en el correo; en el radar actual queda fuera (arribados_bl). En v2 entra a pestaña Salida a DC."},
            {"id": "ESC-08", "titulo": "Marítimo cuya ETA se movió >5 días entre semanas (+9 días)", "hawb_bl": "SINTM-200004", "so": "4229900204",
             "so_local": ["4339900204"], "esperado_correo_hoy": False, "esperado_radar_hoy": False,
             "deberia_v2": "BRECHA HOY: historial_transito no abre versión nueva al cambiar ETA y ni correo ni radar alertan el salto de +9 días. En v2 abre versión y dispara Alerta ETA."},
            {"id": "ESC-09", "titulo": "Línea de stock sin 433 (SINTA-100006 en bi y SINTA-100009 rescatado)", "hawb_bl": "SINTA-100006", "so": "4229900301",
             "so_local": [], "esperado_correo_hoy": True, "esperado_radar_hoy": True,
             "deberia_v2": "Entra a correo y radar aéreo como stock ('Sin pedido en OrderMatch (stock)'), aunque el docstring del correo afirma lo contrario y SINTA-100009 fuerza re-escaneo perpetuo en backfill."},
            {"id": "ESC-10", "titulo": "422 con dos 433", "hawb_bl": "SINTA-100007", "so": "4229900401",
             "so_local": ["4339900401", "4339900402"], "esperado_correo_hoy": True, "esperado_radar_hoy": True,
             "deberia_v2": "Aparece en correo y radar aéreo sin duplicar las 60 unidades de la 422, listando ambos 433 con sus respectivos clientes y Sold-to."},
            {"id": "ESC-11", "titulo": "Embarque con dos MOT (AIR + SEA)", "hawb_bl": "SINTA-100008", "so": "4229900501",
             "so_local": ["4339900501", "4339900502"], "esperado_correo_hoy": True, "esperado_radar_hoy": False,
             "deberia_v2": "BUG HOY: en el correo mot='AIR / SEA' produce sla_hab=NaN y fuera_sla=False pese a llevar >7 días hábiles sin salida; en el radar aéreo se omite por filtrar mot=='AIR'."},
            {"id": "ESC-12", "titulo": "433 con DN en M4 (ingresado SAP) y otro sin DN", "hawb_bl": "SINTM-200005", "so": "4229900601",
             "so_local": ["4339900601", "4339900602"], "esperado_correo_hoy": False, "esperado_radar_hoy": False,
             "deberia_v2": "BRECHA HOY: embarque entregado al DC (>2 d.h.) no aparece en correo ni radar; ningún job cruza M4 para mostrar que 4339900602 (y 4339900105) ya tienen DN en SAP mientras 4339900601 sigue pendiente de ingreso."},
        ],
    }
    (DIR_INSUMOS / "manifiesto_escenarios.json").write_text(
        json.dumps(manifiesto, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    reloj = RelojVirtual(ayer.replace(hour=8, minute=30))
    return EstadoFakes(
        reloj=reloj,
        om_por_dia={1: ruta_om_d1, 2: ruta_om_d2},
        hus_por_dia={1: hus_d1, 2: hus_d2},
        inbox=inbox,
        shipments=shipments,
        enviados=enviados,
        barcos_ais=barcos_ais,
        carpeta_captura_mail=DIR_INSUMOS / "captura_mail",
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera datos sintéticos deterministas en sandbox/")
    parser.add_argument("--hoy", default=FECHA_HOY_DEFAULT, help="Fecha 'hoy' YYYY-MM-DD (default: 2026-11-25)")
    parser.add_argument("--semilla", type=int, default=SEMILLA_DEFAULT, help="Semilla fija determinista")
    args = parser.parse_args()

    estado = construir_escenarios_y_datos(fecha_hoy_str=args.hoy, semilla=args.semilla)
    print(f"Sintéticos generados OK en {SANDBOX}")
    print(f"  Fecha ayer (Día 1): {(datetime.strptime(args.hoy, '%Y-%m-%d') - timedelta(days=1)):%Y-%m-%d} · Fecha hoy (Día 2): {args.hoy}")
    print(f"  OM Día 1 / Día 2: {estado.om_por_dia[1].name}, {estado.om_por_dia[2].name}")
    print(f"  Correos Inbox: {len(estado.inbox)} · Shipments: {len(estado.shipments)} · Enviados semanales: {len(estado.enviados)}")
    print(f"  M4 parquets: {(DIR_M4 / 'm4_latam.parquet').name}, {(DIR_M4 / 'm4_latam_sin_gi.parquet').name}")


if __name__ == "__main__":
    main()
