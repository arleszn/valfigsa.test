#!/usr/bin/env python3
"""
procesar_facturas.py
---------------------
Automatiza el flujo descrito en la prueba:
  1. Extrae datos estructurados de las facturas PDF (folio, cliente, RFC,
     fecha, concepto, subtotal, IVA, total).
  2. Valida que subtotal + IVA == total y que el IVA corresponda al 16%
     del subtotal (detecta errores de captura o de cálculo).
  3. Incorpora correcciones que llegan por correo de un cliente (ej. RFC
     faltante), usando reglas simples de texto (folio + patrón de RFC).
  4. Consolida todo en un Excel con el mismo layout que el registro manual
     original, agregando columnas de validación y observaciones.
  5. Genera un resumen ejecutivo (texto) con incidencias y acciones sugeridas.

Uso:
    python procesar_facturas.py \
        --facturas_dir /ruta/a/facturas \
        --correos_dir /ruta/a/correos \
        --salida /ruta/al/Registro_Consolidado.xlsx

Todas las rutas tienen valores por defecto pensados para correrse tal cual
dentro de esta prueba (ver README.md).
"""
import argparse
import glob
import os
import re
import sys
from datetime import datetime

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. EXTRACCIÓN DE FACTURAS PDF
# ---------------------------------------------------------------------------

FIELD_PATTERNS = {
    "folio": r"Folio:\s*([A-Za-z0-9\-]+)",
    "fecha": r"Fecha:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})",
    "cliente": r"Cliente:\s*(.+)",
    "rfc": r"RFC:\s*(.+)",
    "concepto": r"^\s*([A-Za-zÁÉÍÓÚÑáéíóúñ][A-Za-zÁÉÍÓÚÑáéíóúñ .]+?)\s+\$[0-9,]+\.\d{2}\s*$",
    "subtotal": r"Subtotal:\s*\$([0-9,]+\.\d{2})",
    "iva": r"IVA\s*16%:\s*\$([0-9,]+\.\d{2})",
    "total": r"Total:\s*\$([0-9,]+\.\d{2})",
}


def _to_float(money_str: str) -> float:
    return float(money_str.replace(",", ""))


def extraer_factura(pdf_path: str) -> dict:
    """Extrae los campos clave de una factura PDF usando expresiones regulares
    sobre el texto plano. El formato de las facturas es consistente, por lo
    que un enfoque de regex es suficiente y evita la complejidad/costo de un
    LLM para este paso puramente estructurado."""
    with pdfplumber.open(pdf_path) as pdf:
        texto = "\n".join(page.extract_text() or "" for page in pdf.pages)

    data = {"archivo": os.path.basename(pdf_path)}

    for campo, patron in FIELD_PATTERNS.items():
        m = re.search(patron, texto, re.MULTILINE)
        if not m:
            data[campo] = None
            continue
        valor = m.group(1).strip()
        if campo in ("subtotal", "iva", "total"):
            data[campo] = _to_float(valor)
        else:
            data[campo] = valor

    # Normaliza RFC "NO CAPTURADO" a None para tratarlo como dato faltante
    if data.get("rfc") and "NO CAPTURADO" in data["rfc"].upper():
        data["rfc"] = None

    return data


# ---------------------------------------------------------------------------
# 2. VALIDACIÓN DE CONSISTENCIA
# ---------------------------------------------------------------------------

def validar_factura(f: dict) -> list:
    """Devuelve una lista de incidencias (strings) encontradas en la factura."""
    incidencias = []

    for campo in ("folio", "cliente", "rfc", "fecha", "concepto", "subtotal", "iva", "total"):
        if not f.get(campo):
            incidencias.append(f"Campo faltante: {campo}")

    if f.get("subtotal") is not None and f.get("iva") is not None:
        iva_esperado = round(f["subtotal"] * 0.16, 2)
        if abs(f["iva"] - iva_esperado) > 0.01:
            incidencias.append(
                f"IVA inconsistente: capturado ${f['iva']:.2f}, esperado (16% del subtotal) ${iva_esperado:.2f}"
            )

    if f.get("subtotal") is not None and f.get("iva") is not None and f.get("total") is not None:
        total_esperado = round(f["subtotal"] + f["iva"], 2)
        if abs(f["total"] - total_esperado) > 0.01:
            incidencias.append(
                f"Total inconsistente: capturado ${f['total']:.2f}, esperado (subtotal+IVA) ${total_esperado:.2f}"
            )

    return incidencias


# ---------------------------------------------------------------------------
# 3. INCORPORACIÓN DE CORREOS DE CLIENTES
# ---------------------------------------------------------------------------

def procesar_correos(correos_dir: str) -> dict:
    """Lee correos de texto plano y extrae correcciones/datos por folio.
    Regla simple: busca un folio (F-####) y, en las líneas cercanas, un RFC
    con el patrón estándar mexicano (4 letras/dígitos + 6 dígitos de fecha +
    3 caracteres homoclave)."""
    correcciones = {}
    rfc_pattern = re.compile(r"\b([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})\b")
    folio_pattern = re.compile(r"\b(F-\d{3,5})\b")

    for path in glob.glob(os.path.join(correos_dir, "*.txt")):
        with open(path, encoding="utf-8") as fh:
            texto = fh.read()

        folios = folio_pattern.findall(texto)
        rfcs = rfc_pattern.findall(texto)
        fecha_pago_m = re.search(r"pago.*?(\d{1,2} de \w+ de \d{4})", texto, re.IGNORECASE)

        if folios and rfcs:
            folio = folios[0]
            correcciones.setdefault(folio, {})
            correcciones[folio]["rfc_corregido"] = rfcs[0]
            correcciones[folio]["fuente"] = os.path.basename(path)
            if fecha_pago_m:
                correcciones[folio]["fecha_pago_programada"] = fecha_pago_m.group(1)

    return correcciones


# ---------------------------------------------------------------------------
# 4. CONSOLIDACIÓN + SALIDA A EXCEL
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
ISSUE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
OK_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
BASE_FONT = Font(name="Arial")


def construir_consolidado(facturas: list, correcciones: dict, salida_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Registro consolidado"

    headers = ["Folio", "Cliente", "RFC", "Fecha", "Concepto", "Subtotal",
               "IVA", "Total", "Estatus", "Observaciones"]
    ws.append(headers)
    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    row_num = 2
    for f in sorted(facturas, key=lambda x: x.get("folio") or ""):
        incidencias = list(f["_incidencias"])
        folio = f.get("folio")
        rfc = f.get("rfc")
        obs_extra = []

        # Aplica corrección de correo si existe
        if folio in correcciones:
            corr = correcciones[folio]
            if not rfc and corr.get("rfc_corregido"):
                rfc = corr["rfc_corregido"]
                obs_extra.append(f"RFC actualizado via correo ({corr['fuente']})")
                incidencias = [i for i in incidencias if "rfc" not in i.lower()]
            if corr.get("fecha_pago_programada"):
                obs_extra.append(f"Pago programado: {corr['fecha_pago_programada']}")

        estatus = "Con incidencias" if incidencias else "OK"
        observaciones = "; ".join(incidencias + obs_extra) if (incidencias or obs_extra) else ""

        ws.append([
            folio, f.get("cliente"), rfc, f.get("fecha"), f.get("concepto"),
            f.get("subtotal"), f.get("iva"), f.get("total"), estatus, observaciones,
        ])

        fill = ISSUE_FILL if incidencias else OK_FILL
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = BASE_FONT
            cell.fill = fill
            if col_idx in (6, 7, 8):
                cell.number_format = '$#,##0.00'
        row_num += 1

    # Fila de totales (formula, no valor fijo)
    last_data_row = row_num - 1
    ws.cell(row=row_num, column=5, value="TOTALES").font = Font(name="Arial", bold=True)
    for col_letter in ("F", "G", "H"):
        cell = ws.cell(row=row_num, column=ord(col_letter) - ord("A") + 1)
        cell.value = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
        cell.font = Font(name="Arial", bold=True)
        cell.number_format = '$#,##0.00'

    # Ancho de columnas
    widths = [10, 30, 16, 12, 28, 12, 12, 12, 16, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    wb.save(salida_path)
    return last_data_row


# ---------------------------------------------------------------------------
# 5. RESUMEN EJECUTIVO
# ---------------------------------------------------------------------------

def generar_resumen(facturas: list, correcciones: dict, resumen_path: str):
    total_docs = len(facturas)
    con_incidencias = [f for f in facturas if f["_incidencias"]]
    sin_incidencias = total_docs - len(con_incidencias)
    monto_total = sum(f.get("total") or 0 for f in facturas)

    lineas = []
    lineas.append("RESUMEN EJECUTIVO - PROCESAMIENTO DE FACTURAS")
    lineas.append(f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lineas.append("")
    lineas.append(f"Documentos procesados: {total_docs}")
    lineas.append(f"  - Sin incidencias: {sin_incidencias}")
    lineas.append(f"  - Con incidencias: {len(con_incidencias)}")
    lineas.append(f"Monto total facturado (según campo Total capturado): ${monto_total:,.2f}")
    lineas.append(f"Correos de clientes procesados con correcciones aplicadas: {len(correcciones)}")
    lineas.append("")
    lineas.append("INCIDENCIAS DETECTADAS")
    if not con_incidencias:
        lineas.append("  Ninguna.")
    else:
        alguna = False
        for f in con_incidencias:
            pendientes = [inc for inc in f["_incidencias"]
                          if not ("rfc" in inc.lower() and f.get("folio") in correcciones)]
            if not pendientes:
                continue  # todas las incidencias de esta factura ya se resolvieron con un correo
            alguna = True
            lineas.append(f"  - {f.get('folio')} ({f.get('cliente')}):")
            for inc in pendientes:
                lineas.append(f"      * {inc}")
        if not alguna:
            lineas.append("  Ninguna (todas resueltas con correos de clientes).")
    lineas.append("")
    lineas.append("CORRECCIONES INCORPORADAS DESDE CORREOS DE CLIENTES")
    if not correcciones:
        lineas.append("  Ninguna.")
    else:
        for folio, corr in correcciones.items():
            detalle = f"RFC -> {corr.get('rfc_corregido')}"
            if corr.get("fecha_pago_programada"):
                detalle += f", pago programado {corr['fecha_pago_programada']}"
            lineas.append(f"  - {folio}: {detalle} (fuente: {corr.get('fuente')})")
    lineas.append("")
    lineas.append("ACCIONES SUGERIDAS")
    acciones = []
    for f in con_incidencias:
        folio = f.get("folio")
        pendientes = [i for i in f["_incidencias"]
                      if not (folio in correcciones and "rfc" in i.lower())]
        if any("Total inconsistente" in i for i in pendientes):
            acciones.append(
                f"  - Confirmar con el emisor de {folio} el total correcto: la factura no cuadra "
                f"aritméticamente (subtotal + IVA != total impreso). No pagar hasta aclarar."
            )
        if any("Campo faltante" in i for i in pendientes):
            acciones.append(f"  - Solicitar al emisor de {folio} los datos faltantes antes de archivar.")
    if not acciones:
        acciones.append("  Ninguna acción pendiente; todas las incidencias fueron resueltas.")
    lineas.extend(acciones)
    lineas.append("")
    lineas.append("PROPUESTA DE AUTOMATIZACION (siguientes pasos)")
    lineas.append("  Ver README.md, seccion 'Como escalar esto a produccion'.")

    with open(resumen_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lineas))

    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Procesa facturas PDF y correos, y consolida un registro.")
    parser.add_argument("--facturas_dir", default="/mnt/user-data/uploads")
    parser.add_argument("--correos_dir", default="/mnt/user-data/uploads")
    parser.add_argument("--salida", default="/home/claude/work/data/Registro_Consolidado.xlsx")
    parser.add_argument("--resumen", default="/home/claude/work/data/Resumen_Ejecutivo.txt")
    args = parser.parse_args()

    pdf_paths = sorted(glob.glob(os.path.join(args.facturas_dir, "Factura_*.pdf")))
    if not pdf_paths:
        print("No se encontraron facturas PDF (patrón Factura_*.pdf).", file=sys.stderr)
        sys.exit(1)

    facturas = []
    for path in pdf_paths:
        f = extraer_factura(path)
        f["_incidencias"] = validar_factura(f)
        facturas.append(f)

    correcciones = procesar_correos(args.correos_dir)

    os.makedirs(os.path.dirname(args.salida), exist_ok=True)
    construir_consolidado(facturas, correcciones, args.salida)
    resumen_texto = generar_resumen(facturas, correcciones, args.resumen)

    print(f"OK: {len(facturas)} facturas procesadas -> {args.salida}")
    print(f"OK: resumen ejecutivo -> {args.resumen}")
    print()
    print(resumen_texto)


if __name__ == "__main__":
    main()
