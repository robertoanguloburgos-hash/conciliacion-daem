"""
Conciliador de Cartolas BCI (Excel a Excel - Correlativo Dinámico)
------------------------------------------------------------------
Cruce de alta velocidad basado estrictamente en Monto y Fecha Contable.
Detecta los saltos de hojas del BCI de forma automática mediante bloques secuenciales.
"""

import streamlit as st
import pandas as pd
import openpyxl
import io
import math
from datetime import datetime

# CONFIGURACIÓN DE PÁGINA
st.set_page_config(
    page_title="Conciliador de Cartolas BCI",
    page_icon="🏦",
    layout="wide"
)

def parsear_fecha(texto):
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return None
    if isinstance(texto, datetime):
        return texto
    texto = str(texto).strip().split(" ")[0]
    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None

def limpiar_monto_celda(val):
    if val is None or val == "":
        return None
    if isinstance(val, float) and math.isnan(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    texto = str(val).strip().replace("$", "").replace(".", "")
    texto = texto.replace(",", ".")
    try:
        num = float(texto)
        return None if math.isnan(num) else num
    except ValueError:
        return None

def extraer_movimientos_cartola_convertida(archivo_convertido):
    """
    Lee las filas de la cartola convertida a Excel. Cada vez que encuentra 
    la palabra 'SUCURSAL' o 'FECHA' entiende que cambió de hoja o cartola,
    e incrementa el contador de cartolas de forma dinámica.
    """
    df_raw = pd.read_excel(archivo_convertido)
    movimientos_banco = []
    
    cartola_correlativa = 0
    
    # Índices estándar de columnas por defecto en el formato del BCI
    idx_desc = 2
    idx_doc = 4
    idx_cargo = 6
    idx_abono = 10

    for idx, col_name in enumerate(df_raw.columns):
        val_str = str(col_name).upper()
        if "DESC" in val_str: idx_desc = idx
        elif "DOC" in val_str: idx_doc = idx
        elif "CARGO" in val_str or "CHEQUE" in val_str: idx_cargo = idx
        elif "ABONO" in val_str or "DEPOSIT" in val_str: idx_abono = idx

    for idx, row in df_raw.iterrows():
        # DETECCIÓN DINÁMICA DE CAMBIO DE CARTOLA:
        # Cada vez que vemos la palabra SUCURSAL o FECHA, incrementamos el número de cartola
        row_text_dump = " ".join([str(cell).upper() for cell in row.values if pd.notna(cell)])
        if "SUCURSAL" in row_text_dump or "SALDO DIARIO" in row_text_dump:
            cartola_correlativa += 1
            continue
            
        desc = str(row.iloc[idx_desc]).upper() if pd.notna(row.iloc[idx_desc]) else ""
        if "RESUMEN" in desc or "PERIODO" in desc or "RETENCIONES" in desc:
            continue
            
        fecha_banco = parsear_fecha(row.iloc[0])
        if not fecha_banco:
            continue
            
        raw_doc = row.iloc[idx_doc]
        doc_banco = int(float(str(raw_doc).strip())) if pd.notna(raw_doc) and str(raw_doc).strip().replace(".0","").isdigit() and float(str(raw_doc).strip()) > 0 else None
        
        cargo_banco = limpiar_monto_celda(row.iloc[idx_cargo])
        abono_banco = limpiar_monto_celda(row.iloc[idx_abono])
        
        monto_final = None
        if cargo_banco is not None: monto_final = cargo_banco
        elif abono_banco is not None: monto_final = abono_banco
            
        if monto_final is None:
            continue
            
        movimientos_banco.append({
            "fecha": fecha_banco,
            "documento": doc_banco,
            "monto": abs(monto_final),
            "cartola_calculada": max(1, cartola_correlativa) # Asegurar mínimo 1
        })
        
    return pd.DataFrame(movimientos_banco)

def procesar_excel(excel_control_bytes, df_banco, tolerancia_dias=5):
    wb = openpyxl.load_workbook(io.BytesIO(excel_control_bytes), data_only=False)
    ws = wb.active

    # LOCALIZACIÓN DINÁMICA DE LA FILA DE ENCABEZADOS EN EL CONTROL
    fila_encabezado = None
    columnas = {}
    
    for r in range(1, 20):
        row_vals = [str(cell.value).strip().lower() for cell in ws[r] if cell.value is not None]
        if any("fecha contable" in v or "cargo (-)" in v for v in row_vals):
            fila_encabezado = r
            for cell in ws[r]:
                if cell.value is not None:
                    columnas[str(cell.value).strip()] = cell.column
            break

    if not fila_encabezado:
        raise ValueError("No se pudo encontrar la fila de encabezados reales en tu planilla de control.")

    col_fecha = columnas.get("Fecha contable (*)")
    col_doc = columnas.get("N° documento")
    col_cargo = columnas.get("Cargo (-)")
    col_abono = columnas.get("Abono (+)")
    col_cartola = columnas.get("CARTOLA N°")

    if not col_cartola:
        for k, v in columnas.items():
            if "CARTOLA" in k: col_cartola = v

    used_banco_indices = set()
    emparejados = 0
    sin_match = 0
    total_filas = ws.max_row

    for fila in range(fila_encabezado + 1, total_filas + 1):
        cargo_val = ws.cell(row=fila, column=col_cargo).value if col_cargo else None
        abono_val = ws.cell(row=fila, column=col_abono).value if col_abono else None
        fecha_val = ws.cell(row=fila, column=col_fecha).value if col_fecha else None
        doc_val = ws.cell(row=fila, column=col_doc).value if col_doc else None

        monto_excel = None
        cargo_limpio = limpiar_monto_celda(cargo_val)
        abono_limpio = limpiar_monto_celda(abono_val)
        
        if cargo_limpio is not None: monto_excel = cargo_limpio
        elif abono_limpio is not None: monto_excel = abono_limpio

        if monto_excel is None or monto_excel == 0:
            continue

        fecha_excel = parsear_fecha(fecha_val)
        num_doc_excel = int(float(str(doc_val).strip())) if pd.notna(doc_val) and str(doc_val).strip().replace(".0","").isdigit() and float(str(doc_val).strip()) > 0 else None

        match_idx = None

        # PASO 1: Amarre prioritario por Número de Documento único (Cheques / Transferencias con ID)
        if num_doc_excel is not None and not df_banco.empty:
            cand = df_banco[(df_banco["documento"] == num_doc_excel) & (~df_banco.index.isin(used_banco_indices))]
            if not cand.empty:
                match_idx = cand.index[0]

        # PASO 2: Amarre estricto por Monto Exacto + Ventana de Tiempo (Lógica FIFO secuencial)
        if match_idx is None and not df_banco.empty:
            cand = df_banco[(df_banco["monto"] == round(abs(monto_excel), 2)) & (~df_banco.index.isin(used_banco_indices))]
            for idx_b, fila_b in cand.iterrows():
                if fecha_excel and fila_b["fecha"]:
                    if abs((fecha_excel - fila_b["fecha"]).days) > tolerancia_dias:
                        continue
                match_idx = idx_b
                break

        if match_idx is not None:
            used_banco_indices.add(match_idx)
            cartola_num = df_banco.loc[match_idx, "cartola_calculada"]
            if col_cartola:
                ws.cell(row=fila, column=col_cartola).value = f"CARTOLA {cartola_num}"
            emparejados += 1
        else:
            sin_match += 1

    buffer_out = io.BytesIO()
    wb.save(buffer_out)
    buffer_out.seek(0)
    return buffer_out, emparejados, sin_match

# ENTORNO INTERFAZ
st.title("🏦 Conciliador de Cartolas BCI (Excel a Excel)")
st.markdown("Mapeo cronológico de alta velocidad basado en Monto y Ventana Temporal.")

st.divider()
file_a = st.file_uploader("📊 1. Cargar la Planilla Excel de Control", type=["xlsx"])
file_b = st.file_uploader("📄 2. Cargar la Cartola convertida a Excel", type=["xlsx"])

if file_a and file_b:
    if st.button("🚀 Ejecutar Cruce Contable", type="primary"):
        bytes_a = file_a.read()
        bytes_b = file_b.read()
        
        df_test_a = pd.read_excel(io.BytesIO(bytes_a)).head(15)
        text_dump_a = str(df_test_a.values).upper()
        
        if "SUCURSAL" in text_dump_a or "DEPOSITOS" in text_dump_a:
            bytes_cartola = bytes_a
            bytes_control = bytes_b
        else:
            bytes_cartola = bytes_b
            bytes_control = bytes_a
            
        with st.spinner("Procesando movimientos e identificando saltos de cartola..."):
            df_banco = extraer_movimientos_cartola_convertida(io.BytesIO(bytes_cartola))
            
        if df_banco.empty:
            st.error("No se encontraron registros en el archivo bancario.")
        else:
            st.success(f"✅ Se leyeron {len(df_banco)} movimientos reales en la cartola del banco.")
            
            with st.spinner("Asociando números de cartola en tu archivo..."):
                try:
                    res_buf, ok, nok = procesar_excel(bytes_control, df_banco)
                    st.success(f"🎯 Cruce finalizado: {ok} filas emparejadas con su número de cartola. {nok} filas sin coincidencia.")
                    st.download_button("📥 Descargar Planilla Finalizada", data=res_buf, file_name="control_bancario_conciliado.xlsx")
                except Exception as e:
                    st.error(f"Error durante la escritura: {e}")
