"""
Conciliador de Cartolas BCI (Corregido)
----------------------------------------
Versión adaptada para capturar saltos de línea en el formato oficial BCI
y corregido para iniciar lectura en la Fila 8 del Excel.
"""

import streamlit as st
import pandas as pd
import pdfplumber
import openpyxl
import re
import io
from datetime import datetime

# CONFIGURACIÓN DE PÁGINA
st.set_page_config(
    page_title="Conciliador de Cartolas BCI",
    page_icon="🏦",
    layout="wide"
)

FILA_ENCABEZADO_DEFECTO = 8  # Fila 8 = Donde están los títulos reales en tu Excel

# EXPRESIONES REGULARES CORREGIDAS
# Usamos [\s\n]* para que capture el número aunque esté en la línea de abajo
PATRON_CARTOLA = re.compile(r"CARTOLA\s*N[°ºo][\s\n]*(\d+)", re.IGNORECASE)
PATRON_FECHA = re.compile(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})")
PATRON_MONTO = re.compile(r"\b\d{1,3}(?:\.\d{3})+(?:,\d+)?\b|\b\d+,\d+\b")

def limpiar_monto(texto):
    if texto is None:
        return None
    texto = str(texto).strip()
    if texto in ("", "-", "$"):
        return None
    texto = texto.replace("$", "").strip()
    texto = texto.replace(".", "")   # quita miles
    texto = texto.replace(",", ".")  # decimal
    try:
        return float(texto)
    except ValueError:
        return None

def parsear_fecha(texto):
    if texto is None:
        return None
    if isinstance(texto, datetime):
        return texto
    texto = str(texto).strip()
    formatos = ["%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"]
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None

def extraer_movimientos_pdf(pdf_bytes, progreso_callback=None):
    registros = []
    cartola_actual = None

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_paginas = len(pdf.pages)

        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text() or ""

            # Captura de cartola multi-línea
            coincidencias = PATRON_CARTOLA.findall(texto)
            if coincidencias:
                cartola_actual = int(coincidencias[-1])

            lineas = texto.split("\n")
            for linea in lineas:
                if "CARTOLA N" in linea.upper() and not any(f in linea for f in ["OF VIRT", "OF CENTRA", "PTO MONTT"]):
                    continue

                montos_encontrados = PATRON_MONTO.findall(linea)
                if not montos_encontrados or cartola_actual is None:
                    continue

                fechas_encontradas = PATRON_FECHA.findall(linea)
                fecha_linea = parsear_fecha(fechas_encontradas[0]) if fechas_encontradas else None

                for m in montos_encontrados:
                    valor = limpiar_monto(m)
                    if valor is not None and valor > 100:  # Filtrar números pequeños residuales
                        registros.append({
                            "cartola": cartola_actual,
                            "fecha": fecha_linea,
                            "monto": valor,
                            "texto_linea": linea.strip()
                        })

            if progreso_callback:
                progreso_callback(i + 1, total_paginas)

    return pd.DataFrame(registros)

def encontrar_columnas(ws, fila_encabezado):
    columnas = {}
    for cell in ws[fila_encabezado]:
        if cell.value is not None:
            nombre = str(cell.value).strip()
            columnas[nombre] = cell.column
    return columnas

def procesar_excel(excel_bytes, df_pdf, tolerancia_dias=4, fila_encabezado=FILA_ENCABEZADO_DEFECTO):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
    ws = wb.active

    columnas = encontrar_columnas(ws, fila_encabezado)

    col_fecha = columnas.get("Fecha contable (*)")
    col_cargo = columnas.get("Cargo (-)")
    col_abono = columnas.get("Abono (+)")
    col_cartola = columnas.get("CARTOLA N°")

    faltantes = [nombre for nombre, idx in [
        ("Fecha contable (*)", col_fecha),
        ("Cargo (-)", col_cargo),
        ("Abono (+)", col_abono),
        ("CARTOLA N°", col_cartola),
    ] if idx is None]

    if faltantes:
        raise ValueError(
            f"No se encontraron las columnas en la fila {fila_encabezado}. Columnas leídas: {list(columnas.keys())}"
        )

    pdf_por_monto = {}
    for idx, fila_pdf in df_pdf.iterrows():
        clave = round(fila_pdf["monto"], 2)
        pdf_por_monto.setdefault(clave, []).append(idx)

    used_pdf_indices = set()
    filas_emparejadas = 0
    filas_sin_match = 0
    total_filas = ws.max_row

    for fila in range(fila_encabezado + 1, total_filas + 1):
        cargo_val = ws.cell(row=fila, column=col_cargo).value
        abono_val = ws.cell(row=fila, column=col_abono).value
        fecha_val = ws.cell(row=fila, column=col_fecha).value

        monto_excel = None
        if cargo_val not in (None, "", 0):
            monto_excel = float(cargo_val) if isinstance(cargo_val, (int, float)) else limpiar_monto(cargo_val)
        elif abono_val not in (None, "", 0):
            monto_excel = float(abono_val) if isinstance(abono_val, (int, float)) else limpiar_monto(abono_val)

        if monto_excel is None or monto_excel == 0:
            continue

        fecha_excel = parsear_fecha(fecha_val)
        candidatos = pdf_por_monto.get(round(abs(monto_excel), 2), [])

        match_encontrado = None
        for idx_pdf in candidatos:
            if idx_pdf in used_pdf_indices:
                continue

            fecha_pdf = df_pdf.loc[idx_pdf, "fecha"]
            if fecha_excel is not None and fecha_pdf is not None:
                diferencia_dias = abs((fecha_excel - fecha_pdf).days)
                if diferencia_dias > tolerancia_dias:
                    continue

            match_encontrado = idx_pdf
            break

        if match_encontrado is not None:
            used_pdf_indices.add(match_encontrado)
            numero_cartola = df_pdf.loc[match_encontrado, "cartola"]
            ws.cell(row=fila, column=col_cartola).value = f"CARTOLA {numero_cartola}"
            filas_emparejadas += 1
        else:
            filas_sin_match += 1

    buffer_salida = io.BytesIO()
    wb.save(buffer_salida)
    buffer_salida.seek(0)
    return buffer_salida, filas_emparejadas, filas_sin_match

# INTERFAZ STREAMLIT
st.title("🏦 Conciliador de Cartolas BCI")
st.markdown("Herramienta automatizada de pre-cruce financiero para la visualización y ordenamiento de cartolas.")

st.divider()
col1, col2 = st.columns(2)
with col1:
    archivo_excel = st.file_uploader("📊 Excel de movimientos bancarios (.xlsx)", type=["xlsx"])
with col2:
    archivo_pdf = st.file_uploader("📄 PDF consolidado de Cartolas BCI (.pdf)", type=["pdf"])

with st.expander("⚙️ Opciones avanzadas"):
    fila_encabezado = st.number_input("Fila de encabezados reales", min_value=1, value=FILA_ENCABEZADO_DEFECTO)
    tolerancia = st.slider("Tolerancia de días", min_value=0, max_value=15, value=4)

st.divider()

if archivo_excel and archivo_pdf:
    if st.button("🚀 Procesar y Conciliar", type="primary"):
        pdf_bytes = archivo_pdf.read()
        excel_bytes = archivo_excel.read()
        barra = st.progress(0, text="Extrayendo datos...")

        def actualizar_progreso(act, tot):
            barra.progress(act / tot, text=f"Leyendo página {act} de {tot}...")

        df_pdf = extraer_movimientos_pdf(pdf_bytes, progreso_callback=actualizar_progreso)
        barra.empty()

        if df_pdf.empty:
            st.error("No se detectó información estructurada en el PDF. Inténtalo nuevamente.")
        else:
            st.success(f"✅ Se procesaron {len(df_pdf)} movimientos en {df_pdf['cartola'].nunique()} cartolas.")
            try:
                buf, ok, nok = procesar_excel(excel_bytes, df_pdf, tolerancia_dias=tolerancia, fila_encabezado=int(fila_encabezado))
                st.success(f"🎯 Cruce exitoso: {ok} emparejados, {nok} sin match.")
                st.download_button("⬇️ Descargar Excel Completado", data=buf, file_name="movimientos_bancarios_con_cartola.xlsx")
            except Exception as e:
                st.error(f"Error en procesamiento: {e}")
