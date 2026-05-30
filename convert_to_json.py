"""
convert_to_json.py
==================
Lee Estatus_de_BO.xlsx desde la carpeta data/ y genera data/data.json
para que el dashboard pueda consumirlo sin depender del Excel.

Ejecutado automáticamente por GitHub Actions al subir el Excel.
"""

import pandas as pd
import json
import os
from datetime import datetime

JSON_PATH = os.path.join("data", "data.json")


def es_excel_valido(ruta):
    """Verifica que el archivo sea un Excel válido antes de usarlo."""
    try:
        pd.ExcelFile(ruta, engine='openpyxl')
        return True
    except Exception as e:
        print(f"   ⚠ Archivo inválido ({ruta}): {e}")
        return False


def encontrar_excel():
    print("\n📁 Archivos en el directorio actual:")
    for f in sorted(os.listdir(".")):
        print(f"   {f}")

    carpeta_data = "data"
    if os.path.isdir(carpeta_data):
        print(f"\n📁 Archivos en {carpeta_data}/:")
        for f in sorted(os.listdir(carpeta_data)):
            print(f"   {f}")

    # 1. Buscar por nombres conocidos (más específicos primero)
    nombres_exactos = [
        "Estatus de BO.xlsx", "Estatus_de_BO.xlsx",
        "Estatus BO.xlsx",    "Estatus_BO.xlsx",
        "estatus de bo.xlsx", "estatus_de_bo.xlsx",
        "estatus bo.xlsx",    "estatus_bo.xlsx",
    ]
    carpetas = [carpeta_data, "."]   # buscar primero en data/
    for carpeta in carpetas:
        for nombre in nombres_exactos:
            ruta = os.path.join(carpeta, nombre)
            if os.path.exists(ruta) and es_excel_valido(ruta):
                print(f"\n✅ Excel encontrado: {ruta}")
                return ruta

    # 2. Buscar cualquier .xlsx válido en data/ y raíz
    for carpeta in carpetas:
        base = carpeta if os.path.isdir(carpeta) else "."
        for f in sorted(os.listdir(base)):
            if f.lower().endswith(".xlsx"):
                ruta = os.path.join(base, f)
                if es_excel_valido(ruta):
                    print(f"\n✅ Excel encontrado (búsqueda amplia): {ruta}")
                    return ruta
                else:
                    print(f"   ⏭ Saltando archivo inválido: {ruta}")

    raise FileNotFoundError(
        "No se encontró ningún archivo .xlsx en el repositorio.\n"
        "Asegúrate de haber subido el Excel al repo."
    )


def safe_str(val):
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    return str(val).strip()


def safe_date(val):
    if pd.isna(val) if not isinstance(val, str) else False:
        return ""
    try:
        if isinstance(val, (pd.Timestamp, datetime)):
            return val.strftime("%Y-%m-%d")
        d = pd.to_datetime(val, errors="coerce")
        return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""
    except Exception:
        return ""


def safe_num(val):
    try:
        f = float(val)
        return round(f, 4) if not (f != f) else 0   # NaN check
    except Exception:
        return 0


def main():
    EXCEL_PATH = encontrar_excel()
    if not EXCEL_PATH:
        raise FileNotFoundError("No se encontró ningún archivo Excel válido en el repositorio.")

    print(f"📂 Leyendo: {EXCEL_PATH}")
    xl = pd.ExcelFile(EXCEL_PATH, engine='openpyxl')
    print(f"   Hojas encontradas: {xl.sheet_names}")

    # ── Hoja 1 (índice 1): " Ordenes de Compra" → ÓRDENES ──────────────────
    df1 = pd.read_excel(EXCEL_PATH, sheet_name=1, dtype=str, engine='openpyxl')
    df1 = df1.fillna("")

    # ── Hoja 3 (índice 3): "RP01-23 Ordenes de Compra" → Estado OC por folio
    estado_oc_map = {}
    try:
        df_rp = pd.read_excel(EXCEL_PATH, sheet_name=3, dtype=str, engine='openpyxl').fillna("")
        if 'Estado OC' in df_rp.columns and 'Folio' in df_rp.columns:
            for _, r in df_rp.iterrows():
                folio = safe_str(r.get('Folio', ''))
                estado = safe_str(r.get('Estado OC', '')).upper()
                if folio:
                    estado_oc_map[folio] = estado
            print(f"   ✅ Estado OC cargado: {len(estado_oc_map)} folios")
    except Exception as e:
        print(f"   ⚠ No se pudo leer Estado OC: {e}")

    orders = []
    for _, row in df1.iterrows():
        folio = safe_str(row.get("Folio", ""))
        orders.append({
            "Folio":              folio,
            "Sucursal":           safe_str(row.get("Sucursal", "")),
            "Fecha":              safe_str(row.get("Fecha", "")),
            "Proveedor":          safe_str(row.get("Proveedor", "")),
            "Plazo Pago":         safe_str(row.get("Plazo Pago", "")),
            "Subtotal":           safe_num(row.get("Subtotal", 0)),
            "Importe IVA":        safe_num(row.get("Importe IVA", 0)),
            "Importe":            safe_num(row.get("Importe", 0)),
            "Flete x Cobrar":     safe_str(row.get("Flete x Cobrar", "")),
            "Transportista":      safe_str(row.get("Transportista", "")),
            "Fecha Entrega":      safe_str(row.get("Fecha Entrega", "")),
            "Acepta Parciales":   safe_str(row.get("Acepta Parciales", "")),
            "Observaciones":      safe_str(row.get("Observaciones", "")),
            "Estatus":            safe_str(row.get("Estatus", "")),
            "Back Order":         safe_str(row.get("Back Order", "")),
            "Referencia Lista":   safe_str(row.get("Referencia Lista", "")),
            "Comprador Capturo":  safe_str(row.get("Comprador Capturo", "")),
            "Estado OC":          estado_oc_map.get(folio, "RECIBIDO"),
        })

    print(f"   ✅ Órdenes procesadas: {len(orders)}")

    # ── Hoja 0 (índice 0): "Productos Incluidos en" → PRODUCTOS ────────────
    df2 = pd.read_excel(EXCEL_PATH, sheet_name=0, dtype=str, engine='openpyxl')
    df2 = df2.fillna("")

    products = {}
    for _, row in df2.iterrows():
        folio = safe_str(row.get("FOLIO OC", ""))
        if not folio:
            continue
        if folio not in products:
            products[folio] = []
        products[folio].append({
            "FOLIO OC":       folio,
            "FECHA OC":       safe_str(row.get("FECHA OC", "")),
            "PROVEEDOR":      safe_str(row.get("PROVEEDOR", "")),
            "CLAVE":          safe_str(row.get("CLAVE", "")),
            "DESCRIPCION":    safe_str(row.get("DESCRIPCION", "")),
            "MARCA":          safe_str(row.get("MARCA", "")),
            "UNIDAD":         safe_str(row.get("UNIDAD", "")),
            "CONTENIDO":      safe_str(row.get("CONTENIDO", "")),
            "PZASXUNI":       safe_num(row.get("PZASXUNI", 0)),
            "CANTIDAD PED":   safe_num(row.get("CANTIDAD PED", 0)),
            "CANTIDAD PEND":  safe_num(row.get("CANTIDAD PEND", 0)),
            "BACKORDER":      safe_num(row.get("BACKORDER", 0)),
            "FECHA ENTREGA":  safe_str(row.get("FECHA ENTREGA", "")),
            "ESTATUS COMPRA": safe_str(row.get("ESTATUS COMPRA", "")),
            "COMPRADOR":      safe_str(row.get("COMPRADOR", "")),
        })

    total_prods = sum(len(v) for v in products.values())
    print(f"   ✅ Claves procesadas:  {total_prods}  (en {len(products)} folios)")

    # ── Generar JSON ──────────────────────────────────────────────────────
    output = {
        "generated":     datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "totalOrders":   len(orders),
        "totalProducts": total_prods,
        "orders":        orders,
        "products":      products,
    }

    os.makedirs("data", exist_ok=True)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = os.path.getsize(JSON_PATH) / 1024
    print(f"\n✅ Generado: {JSON_PATH}  ({size_kb:.1f} KB)")
    print(f"   Generado: {output['generated']}")


if __name__ == "__main__":
    main()
