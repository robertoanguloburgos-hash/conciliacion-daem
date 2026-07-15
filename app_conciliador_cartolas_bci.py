"""
Conciliador de Cartolas BCI (Versión de Máxima Precisión)
--------------------------------------------------------
Corrige el error de congelamiento de Cartola mediante Regex de proximidad de caracteres.
Soporta carga masiva de múltiples PDFs mensuales en simultáneo.
"""

import streamlit as st
import pandas as pd
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

FILA_ENCABEZADO_DEFECTO = 8  # Fila estándar de tus columnas en el Excel

# EXPRESIONES REGULARES DE ALTA TOLERANCIA
# Busca la palabra CARTOLA y captura el primer número que encuentre en los siguientes 20 caracteres
PATRON_CARTOLA = re.compile(r"CARTOLA[\s\S]{0,20}?(\d+)", re.IGNORECASE)
PATRON_FECHA = re.compile(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})")
PATRON_MONTO = re.compile(r"\b\d{1,3}(?:\.\d{3})+(?:,\d+)?\b|\b\d+(?:,\d+)?\b")

try:
    import pdfplumber
except ImportError:
    st.error("Falta instalar la librería 'pdfplumber'. Verifica tu archivo requirements.txt")

def limpiar_monto(texto):
    if texto is None:
        return None
    texto = str(texto).strip()
    if texto in ("", "-", "$"):
        return None
    texto = texto.replace("$", "").strip()
    texto = texto.replace(".", "")   
    texto = texto.replace(",", ".")  
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

def extraer_datos_de_lista_pdfs(lista_archivos_pdf, progreso_bar):
    registros = []
    total_archivos = len(lista_archivos_pdf)
    
    for idx_arch, archivo in enumerate(lista_archivos_pdf):
        cartola_actual = None
        pdf_bytes = archivo.read()
        
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto = pagina.extract_text() or ""

                # LÓGICA ROBUSTA: Captura el número de cartola real saltándose las cajas de texto del BCI
                coincidencias = PATRON_CARTOLA.findall(texto)
                if coincidencias:
                    cartola_actual = int(coincidencias[-1])

                lineas = texto.split("\n")
                for linea in lineas:
                    if "CASILLA 136-D" in linea or "GERENCIA DE CLIENTES" in linea:
                        continue

                    montos_encontrados = PATRON_MONTO.findall(linea)
                    if not montos_encontrados:
                        continue

                    fechas_encontradas = PATRON_FECHA.findall(linea)
                    fecha_linea = parsear_fecha(fechas_encontradas[0]) if fechas_encontradas else None

                    # Rastrear folios largos únicos
                    doc_match = re.search(r"\b(96\d{7}|\d{7})\b", linea)
                    num_doc_pdf = int(doc_match.group(1)) if doc_match else None

                    # Rescatar el Saldo Diario de control de la línea
                    saldo_diario_banco = None
                    if len(montos_encontrados) >= 2:
                        saldo_diario_banco = limpiar_monto(montos_encontrados[-1])

                    for m in montos_encontrados[:-1] if len(montos_encontrados) > 1 else montos_encontrados:
                        valor = limpiar_monto(m)
                        if valor is not None and valor > 100 and valor not in [2025, 2026]:
                            registros.append({
                                "cartola": cartola_actual if cartola_actual is not None else 1,
                                "fecha": fecha_linea,
                                "monto": valor,
                                "documento": num_doc_pdf,
                                "saldo_diario": saldo_diario_banco,
                                "texto_linea": linea.strip()
                            })
                            
        progreso_bar.progress((idx_arch + 1) / total_archivos, text=f"Procesando: {archivo.name}")
        
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
    col_doc = columnas.get("N° documento")
    col_cargo = columnas.get("Cargo (-)")
    col_abono = columnas.get("Abono (+)")
    col_saldo = columnas.get("SALDOS")
    col_cartola = columnas.get("CARTOLA N°")

    if not col_cartola:
        for k, v in columnas.items():
            if "CARTOLA" in k: col_cartola = v

    faltantes = [n for n, idx in [
        ("Fecha contable (*)", col_fecha), ("Cargo (-)", col_cargo), 
        ("Abono (+)", col_abono), ("CARTOLA N°", col_cartola)
    ] if idx is None]

    if faltantes:
        raise ValueError(f"Faltan columnas requeridas en la Fila {fila_encabezado}: {faltantes}")

    used_pdf_indices = set()
    filas_emparejadas = 0
    filas_sin_match = 0
    total_filas = ws.max_row

    for fila in range(fila_encabezado + 1, total_filas + 1):
        cargo_val = ws.cell(row=fila, column=col_cargo).value
        abono_val = ws.cell(row=fila, column=col_abono).value
        fecha_val = ws.cell(row=fila, column=col_fecha).value
        doc_val = ws.cell(row=fila, column=col_doc).value if col_doc else None
        saldo_val = ws.cell(row=fila, column=col_saldo).value if col_saldo else None

        monto_excel = None
        if cargo_val not in (None, "", 0):
            monto_excel = float(cargo_val) if isinstance(cargo_val, (int, float)) else limpiar_monto(cargo_val)
        elif abono_val not in (None, "", 0):
            monto_excel = float(abono_val) if isinstance(abono_val, (int, float)) else limpiar_monto(abono_val)

        if monto_excel is None or monto_excel == 0:
            continue

        fecha_excel = parsear_fecha(fecha_val)
        num_doc_excel = int(doc_val) if isinstance(doc_val, (int, float)) and doc_val > 0 else None
        val_saldo_excel = float(saldo_val) if isinstance(saldo_val, (int, float)) else limpiar_monto(saldo_val)

        match_encontrado = None

        # CRUCE DE PRECISIÓN 1: Por número de documento único
        if num_doc_excel is not None:
            candidatos_doc = df_pdf[(df_pdf["documento"] == num_doc_excel) & (~df_pdf.index.isin(used_pdf_indices))]
            if not candidatos_doc.empty:
                match_encontrado = candidatos_doc.index[0]

        # CRUCE DE PRECISIÓN 2: Por Monto + Ventana de Tiempo + Validación del Saldo Diario
        if match_encontrado is None:
            candidatos_monto = df_pdf[(df_pdf["monto"] == round(abs(monto_excel), 2)) & (~df_pdf.index.isin(used_pdf_indices))]
            
            for idx_pdf, fila_pdf in candidatos_monto.iterrows():
                fecha_pdf = fila_pdf["fecha"]
                if fecha_excel is not None and fecha_pdf is not None:
                    if abs((fecha_excel - fecha_pdf).days) > tolerancia_dias:
                        continue
                
                # Desempate matemático estricto usando el saldo diario acumulado
                if val_saldo_excel is not None and fila_pdf["saldo_diario"] is not None:
                    if round(val_saldo_excel, 0) != round(fila_pdf["saldo_diario"], 0):
                        continue

                match_encontrado = idx_pdf
                break

        # Fallback de seguridad (FIFO) si el documento es 0
        if match_encontrado is None and num_doc_excel is None:
            candidatos_fallback = df_pdf[(df_pdf["monto"] == round(abs(monto_excel), 2)) & (~df_pdf.index.isin(used_pdf_indices))]
            if not candidatos_fallback.empty:
                match_encontrado = candidatos_fallback.index[0]

        if match_encontrado is not None:
            used_pdf_indices.add(match_encontrado)
            num_cartola = df_pdf.loc[match_encontrado, "cartola"]
            ws.cell(row=fila, column=col_cartola).value = f"CARTOLA {num_cartola}"
            filas_emparejadas += 1
        else:
            filas_sin_match += 1

    buffer_salida = io.BytesIO()
    wb.save(buffer_salida)
    buffer_salida.seek(0)
    return buffer_salida, filas_emparejadas, filas_sin_match

# INTERFAZ STREAMLIT
st.title("🏦 Conciliador de Cartolas BCI (Multimes)")
st.markdown("Sube tu Excel y los archivos PDF correspondientes a cada mes de forma simultánea.")

st.divider()
archivo_excel = st.file_uploader("Subir la Planilla Excel de Control", type=["xlsx"])
archivos_pdf = st.file_uploader("Subir Cartolas PDF (Puedes seleccionar los 6 meses juntos)", type=["pdf"], accept_multiple_files=True)

with st.expander("⚙️ Ajustes finos"):
    fila_encabezado = st.number_input("Fila donde dice 'Fecha transacción'", min_value=1, value=FILA_ENCABEZADO_DEFECTO)
    tolerancia = st.slider("Ventana de desfase de fechas (Días)", min_value=0, max_value=20, value=4)

st.divider()

if archivo_excel and archivos_pdf:
    if st.button("🚀 Ejecutar Conciliación", type="primary"):
        excel_bytes = archivo_excel.read()
        barra_progreso = st.progress(0, text="Cargando documentos...")
        
        df_pdf = extraer_datos_de_lista_pdfs(archivos_pdf, barra_progreso)
        barra_progreso.empty()

        if df_pdf.empty:
            st.error("Error al procesar los documentos PDF.")
        else:
            st.success(f"✅ Éxito: Se consolidaron {len(df_pdf)} movimientos desde los PDF bancarios.")
            try:
                buf, ok, nok = procesar_excel(excel_bytes, df_pdf, tolerancia_dias=tolerancia, fila_encabezado=int(fila_encabezado))
                st.success(f"🎯 Cruce finalizado con éxito: {ok} movimientos identificados con su respectiva Cartola. {nok} filas sin match.")
                st.download_button("📥 Descargar Excel Completado", data=buf, file_name="movimientos_bancarios_con_cartola.xlsx")
            except Exception as e:
                st.error(f"Error al procesar las celdas del Excel: {e}")
