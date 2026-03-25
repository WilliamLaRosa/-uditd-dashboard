#!/usr/bin/env python3
"""
excel_to_json.py
════════════════════════════════════════════════════════
Convierte SEGUIM_BIENES_SERVICIOS_UDITD.xlsx → datos_bienes.json
automáticamente.

USO:
  python excel_to_json.py
  python excel_to_json.py mi_archivo.xlsx

El script busca el Excel en la misma carpeta. Si no se pasa
argumento, usa el nombre por defecto definido en EXCEL_FILE.
════════════════════════════════════════════════════════
"""

import sys
import json
import os
from datetime import datetime, date
import pandas as pd

# ── CONFIGURACIÓN ──────────────────────────────────────
EXCEL_FILE  = "SEGUIM_BIENES_SERVICIOS_UDITD.xlsx"   # nombre del Excel fuente
OUTPUT_FILE = "datos_bienes.json"                     # JSON de salida
SHEET_NAME  = "Hoja1"                                 # hoja del Excel
HEADER_ROW  = 1    # fila 0-indexada donde están los encabezados (fila 2 en Excel)
DATA_START  = 2    # primera fila de datos (fila 3 en Excel)
DATA_END    = 30   # última fila de datos (exclusivo) — ajustar si crece la tabla
# ──────────────────────────────────────────────────────


def to_safe(v):
    """Convierte cualquier valor a un tipo seguro para JSON."""
    if v is None:
        return None
    if isinstance(v, (datetime, date, pd.Timestamp)):
        try:
            return v.strftime("%Y-%m-%d")
        except Exception:
            return str(v)
    if isinstance(v, float):
        if pd.isna(v):
            return None
        return int(v) if v == int(v) else round(v, 2)
    if isinstance(v, int):
        return v
    if isinstance(v, str):
        s = v.strip()
        return s if s else None
    return None


def abrev_area(nombre):
    """Devuelve una etiqueta corta para el área."""
    if not nombre:
        return "SIN ÁREA"
    n = nombre.upper()
    if "CIRUGIA" in n or "CIRUGÍA" in n:
        return "CIRUGIA EXP."
    if "INVESTIGACION" in n or "INVESTIGACIÓN" in n:
        return "SUIIT"
    if "NORMALIZACION" in n or "NORMALIZACIÓN" in n or "DOCENCIA" in n:
        return "SUNTDD"
    return nombre[:25]


def procesar(excel_path):
    print(f"  Leyendo: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, header=None)

    cols = [
        "expediente", "fecha_req", "area_usuaria", "asunto",
        "tipo", "presupuesto_est", "nro_orden", "siaf",
        "fecha_contrato", "monto", "dependencia", "usuario",
        "fecha_atencion", "estado", "comentario"
    ]

    # Detectar el total real de filas con datos
    end = DATA_END
    for i in range(DATA_START, len(df)):
        val = df.iloc[i, 0]
        if isinstance(val, float) and pd.isna(val):
            if all(pd.isna(df.iloc[i, j]) for j in range(5)):
                end = i
                break

    data_df = df.iloc[DATA_START:end].copy()
    data_df.columns = cols[: len(data_df.columns)]
    # Rellenar columnas faltantes
    for c in cols:
        if c not in data_df.columns:
            data_df[c] = None
    data_df = data_df[cols].reset_index(drop=True)

    registros = []
    for _, row in data_df.iterrows():
        r = {c: to_safe(row[c]) for c in cols}
        r["area_corta"] = abrev_area(r.get("area_usuaria"))
        # normalizar estado
        estado = (r.get("estado") or "").strip().upper()
        if estado not in ("ATENDIDO", "EN PROCESO", "PENDIENTE"):
            estado = "SIN ESTADO"
        r["estado"] = estado
        # normalizar tipo
        tipo = (r.get("tipo") or "").strip().upper()
        if tipo not in ("BIEN", "SERVICIO"):
            tipo = None
        r["tipo"] = tipo
        registros.append(r)

    # ── RESUMEN ESTADÍSTICO ──────────────────────────────
    total = len(registros)
    atendidos   = sum(1 for r in registros if r["estado"] == "ATENDIDO")
    en_proceso  = sum(1 for r in registros if r["estado"] == "EN PROCESO")
    pendientes  = sum(1 for r in registros if r["estado"] == "PENDIENTE")
    sin_estado  = total - atendidos - en_proceso - pendientes

    ppto_total  = sum(r["presupuesto_est"] for r in registros if r["presupuesto_est"])
    monto_total = sum(r["monto"] for r in registros if r["monto"])
    ppto_pen    = ppto_total - monto_total   # presupuesto sin contratar

    # por área
    areas = {}
    for r in registros:
        a = r["area_corta"]
        if a not in areas:
            areas[a] = {"total": 0, "atendidos": 0, "en_proceso": 0,
                        "ppto": 0, "monto": 0}
        areas[a]["total"] += 1
        if r["estado"] == "ATENDIDO":
            areas[a]["atendidos"] += 1
        elif r["estado"] == "EN PROCESO":
            areas[a]["en_proceso"] += 1
        if r["presupuesto_est"]:
            areas[a]["ppto"] += r["presupuesto_est"]
        if r["monto"]:
            areas[a]["monto"] += r["monto"]

    # por tipo
    bienes    = sum(1 for r in registros if r["tipo"] == "BIEN")
    servicios = sum(1 for r in registros if r["tipo"] == "SERVICIO")

    # timeline mensual
    meses_label = ["ENE","FEB","MAR","ABR","MAY","JUN",
                   "JUL","AGO","SET","OCT","NOV","DIC"]
    reqs_mes = [0]*12
    montos_mes = [0.0]*12
    for r in registros:
        f = r.get("fecha_req")
        if f and len(str(f)) >= 7:
            try:
                m = int(str(f)[5:7]) - 1
                reqs_mes[m] += 1
            except Exception:
                pass
        fc = r.get("fecha_contrato")
        if fc and len(str(fc)) >= 7 and r.get("monto"):
            try:
                m = int(str(fc)[5:7]) - 1
                montos_mes[m] += r["monto"]
            except Exception:
                pass

    resultado = {
        "meta": {
            "archivo_fuente": os.path.basename(excel_path),
            "ultima_actualizacion": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "total_registros": total,
        },
        "resumen": {
            "total": total,
            "atendidos": atendidos,
            "en_proceso": en_proceso,
            "pendientes": pendientes,
            "sin_estado": sin_estado,
            "ppto_estimado_total": round(ppto_total, 2),
            "monto_contratado_total": round(monto_total, 2),
            "ppto_pendiente": round(ppto_pen, 2),
            "pct_ejecucion": round(monto_total / ppto_total * 100, 1) if ppto_total else 0,
            "bienes": bienes,
            "servicios": servicios,
        },
        "por_area": areas,
        "timeline": {
            "meses": meses_label,
            "requerimientos": reqs_mes,
            "montos_contratados": [round(v, 2) for v in montos_mes],
        },
        "registros": registros,
    }

    return resultado


def main():
    excel_path = sys.argv[1] if len(sys.argv) > 1 else EXCEL_FILE

    if not os.path.exists(excel_path):
        print(f"ERROR: No se encontró el archivo '{excel_path}'")
        print(f"  Coloca el Excel en la misma carpeta que este script.")
        sys.exit(1)

    print("=" * 52)
    print("  Conversor Excel → JSON · UDITD Bienes y Servicios")
    print("=" * 52)

    resultado = procesar(excel_path)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    r = resultado["resumen"]
    print(f"\n  ✓ JSON generado: {OUTPUT_FILE}")
    print(f"  ✓ Registros procesados : {r['total']}")
    print(f"  ✓ Atendidos            : {r['atendidos']}")
    print(f"  ✓ En proceso           : {r['en_proceso']}")
    print(f"  ✓ Ppto estimado        : S/ {r['ppto_estimado_total']:,.2f}")
    print(f"  ✓ Monto contratado     : S/ {r['monto_contratado_total']:,.2f}")
    print(f"  ✓ % Ejecución          : {r['pct_ejecucion']}%")
    print()
    print("  Sube datos_bienes.json a GitHub y el dashboard")
    print("  se actualizará automáticamente en ~1 minuto.")
    print("=" * 52)


if __name__ == "__main__":
    main()
