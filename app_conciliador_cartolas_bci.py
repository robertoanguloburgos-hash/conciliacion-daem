"""
Conciliador de Cartolas BCI (Versión Libre de Errores de Carga)
-------------------------------------------------------------
Detecta de forma dinámica las filas y columnas correctas sin importar el orden
de los casilleros ni la estructura rígida de filas.
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

def mapear_cartola_por_fecha(fecha_contable):
    if not fecha_contable:
        return 1
    dia = fecha_contable.day
    mes = fecha_contable.month
    
    if mes == 1:
        if dia <= 2: return 1
        elif dia <= 5: return 2
        elif dia <= 6: return 3
        elif dia <= 7: return 4
        elif dia <= 9: return 5
        elif dia <= 12: return 6
        elif dia <= 13: return 7
        elif dia <= 14: return 8
        elif dia <= 19: return 9
        elif dia <= 20: return 10
        elif dia <= 23: return 11
        elif dia <= 26: return 12
        elif dia <= 28: return 13
        elif dia <= 29: return 14
        else: return 15
    return mes + 14  

def extraer_movimientos_cartola_convertida(archivo_convertido):
    df_raw = pd.read_excel(archivo_convertido)
    movimientos_banco = []
    
    # Buscar dinámicamente cuál es la columna de descripción y montos
    idx_desc = 2
    idx_doc = 4
    idx_cargo = 6
    idx_abono = 10
    idx_saldo = 12

    for idx, col_name in enumerate(df_raw.columns):
        val_str = str(col_name).upper()
        if "DESC" in val_str: idx_desc = idx
        elif "DOC" in val_str: idx_doc = idx
        elif "CARGO" in val_str or "CHEQUE" in val_str: idx_cargo = idx
        elif "ABONO" in val_str or "DEPOSIT" in val_str: idx_abono = idx
        elif "SALDO" in val_str: idx_saldo = idx

    for idx, row in df_raw.iterrows():
        desc = str(row.iloc[idx_desc]).upper() if pd.notna(row.iloc[idx_desc]) else ""
        if "RESUMEN" in desc or "PERIODO" in desc or "RETENCIONES" in desc or "FECHA" in desc:
            continue
            
        fecha_banco = parsear_fecha(row.iloc[0])
        if not fecha_banco:
            continue
            
        raw_doc = row.iloc[idx_doc]
        doc_banco = int(float(str(raw_doc).strip())) if pd.notna(raw_doc) and str(raw_doc).strip().replace(".0","").isdigit() and float(str(raw_doc).strip()) > 0 else None
        
        cargo_banco = limpiar_monto_celda(row.iloc[idx_cargo])
        abono_banco = limpiar_monto_celda(row.iloc[idx_abono])
        saldo_banco = limpiar_monto_celda(row.iloc[idx_saldo])
        
        monto_final = None
        if cargo_banco is not None: monto_final = cargo_banco
        elif abono_banco is not None: monto_final = abono_banco
            
        if monto_final is None:
            continue
            
        movimientos_banco.append({
            "fecha": fecha_banco,
            "documento": doc_banco,
            "monto": abs(monto_final),
            "saldo_diario": saldo_banco,
            "cartola_calculada": mapear_cartola_por_fecha(fecha_banco)
        })
        
    return pd.DataFrame(movimientos_banco)

def procesar_excel(excel_control_bytes, df_banco):
    wb = openpyxl.load_workbook(io.BytesIO(excel_control_bytes), data_only=False)
    ws = wb.active

    # ESCANEO DINÁMICO: Buscar en qué fila están las columnas reales del control
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
        raise ValueError("No se pudo encontrar la fila de encabezados en la planilla de control. Verifica que existan las columnas 'Fecha contable (*)' y 'Cargo (-)'")

    col_fecha = columnas.get("Fecha contable (*)")
    col_doc = columnas.get("N° documento")
    col_cargo = columnas.get("Cargo (-)")
    col_abono = columnas.get("Abono (+)")
    col_saldo = columnas.get("SALDOS")
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
        saldo_val = ws.cell(row=fila, column=col_saldo).value if col_saldo else None

        monto_excel = None
        cargo_limpio = limpiar_monto_celda(cargo_val)
        abono_limpio = limpiar_monto_celda(abono_val)
        
        if cargo_limpio is not None: monto_excel = cargo_limpio
        elif abono_limpio is not None: monto_excel = abono_limpio

        if monto_excel is None or monto_excel == 0:
            continue

        fecha_excel = parsear_fecha(fecha_val)
        num_doc_excel = int(float(str(doc_val).strip())) if pd.notna(doc_val) and str(doc_val).strip().replace(".0","").isdigit() and float(str(doc_val).strip()) > 0 else None
        saldo_excel = float(saldo_val) if isinstance(saldo_val, (int, float)) else limpiar_monto_celda(saldo_val)

        match_idx = None

        # CRITERIO 1: Cruce por N° Documento exacto
        if num_doc_excel is not None and not df_banco.empty:
            cand = df_banco[(df_banco["documento"] == num_doc_excel) & (~df_banco.index.isin(used_banco_indices))]
            if not cand.empty:
                match_idx = cand.index[0]

        # CRITERIO 2: Cruce por combinación de Monto + Saldo de Control Bancario
        if match_idx is None and not df_banco.empty:
            cand = df_banco[(df_banco["monto"] == round(abs(monto_excel), 2)) & (~df_banco.index.isin(used_banco_indices))]
            for idx_b, fila_b in cand.iterrows():
                if fecha_excel and fila_b["fecha"]:
                    if abs((fecha_excel - fila_b["fecha"]).days) > 5:
                        continue
                if saldo_excel and fila_b["saldo_diario"]:
                    if round(saldo_excel, 0) != round(fila_b["saldo_diario"], 0):
                        continue
                match_idx = idx_b
                break

        # Fallback de cola (FIFO) secuencial por monto bruto
        if match_idx is None and not df_banco.empty:
            cand = df_banco[(df_banco["monto"] == round(abs(monto_excel), 2)) & (~df_banco.index.isin(used_banco_indices))]
            if not cand.empty:
                match_idx = cand.index[0]

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
st.markdown("Procesa las cartolas en formato planilla. El orden de los casilleros se detecta de forma dinámica.")

st.divider()
file_a = st.file_uploader("📊 Cargar Archivo A (.xlsx)", type=["xlsx"])
file_b = st.file_uploader("📄 Cargar Archivo B (.xlsx)", type=["xlsx"])

if file_a and file_b:
    if st.button("🚀 Iniciar Cruce Automatizado", type="primary"):
        bytes_a = file_a.read()
        bytes_b = file_b.read()
        
        # Identificar inteligentemente cuál archivo es la cartola y cuál es el control
        df_test_a = pd.read_excel(io.BytesIO(bytes_a)).head(15)
        text_dump_a = str(df_test_a.values).upper()
        
        if "SUCURSAL" in text_dump_a or "DEPOSITOS" in text_dump_a:
            bytes_cartola = bytes_a
            bytes_control = bytes_b
        else:
            bytes_cartola = bytes_b
            bytes_control = bytes_a
            
        with st.spinner("Mapeando transacciones de la cartola bancaria..."):
            df_banco = extraer_movimientos_cartola_convertida(io.BytesIO(bytes_cartola))
            
        if df_banco.empty:
            st.error("No se detectaron filas procesables en la cartola bancaria.")
        else:
            st.success(f"✅ Se estructuraron {len(df_banco)} movimientos desde la cartola.")
            
            with st.spinner("Rellenando las celdas vacías del control..."):
                try:
                    res_buf, ok, nok = procesar_excel(bytes_control, df_banco)
                    st.success(f"🎯 Cruce finalizado: {ok} filas emparejadas con su número de cartola. {nok} filas sin match.")
                    st.download_button("📥 Descargar Planilla Finalizada", data=res_buf, file_name="control_bancario_conciliado.xlsx")
                except Exception as e:
                    st.error(f"Error durante el llenado: {e}")
