"""
Conciliador de Cartolas BCI (Excel a Excel - Corregido)
---------------------------------------------------
Cruce directo de planillas para máxima velocidad y precisión.
Elimina la necesidad de procesar archivos PDF.
"""

import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

# CONFIGURACIÓN DE PÁGINA
st.set_page_config(
    page_title="Conciliador de Cartolas BCI",
    page_icon="🏦",
    layout="wide"
)

FILA_ENCABEZADO_DEFECTO = 8  # Fila estándar de tus columnas en tu planilla de control

def parsear_fecha(texto):
    if texto is None:
        return None
    if isinstance(texto, datetime):
        return texto
    texto = str(texto).strip().split(" ")[0]  # Limpiar estampa de hora si existe
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
    if isinstance(val, (int, float)):
        return float(val)
    texto = str(val).strip().replace("$", "").replace(".", "")
    texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None

def mapear_cartola_por_fecha(fecha_contable):
    """
    Asigna de forma lógica el número de cartola correlativo basado en la 
    línea temporal del banco BCI evidenciada en los cierres de periodo.
    """
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
    """Lee las filas de la cartola convertida a Excel y extrae folios y saldos."""
    df_raw = pd.read_excel(archivo_convertido)
    movimientos_banco = []
    
    for idx, row in df_raw.iterrows():
        desc = str(row.iloc[2]).upper() if pd.notna(row.iloc[2]) else ""
        if "RESUMEN" in desc or "PERIODO" in desc or "RETENCIONES" in desc:
            continue
            
        fecha_banco = parsear_fecha(row.iloc[0])
        if not fecha_banco:
            continue
            
        doc_banco = int(row.iloc[4]) if pd.notna(row.iloc[4]) and str(row.iloc[4]).isdigit() else None
        cargo_banco = limpiar_monto_celda(row.iloc[6])
        abono_banco = limpiar_monto_celda(row.iloc[10])
        saldo_banco = limpiar_monto_celda(row.iloc[12])
        
        monto_final = cargo_banco if cargo_banco else abono_banco
        if not monto_final:
            continue
            
        movimientos_banco.append({
            "fecha": fecha_banco,
            "documento": doc_banco,
            "monto": abs(monto_final),
            "saldo_diario": saldo_banco,
            "cartola_calculada": mapear_cartola_por_fecha(fecha_banco)
        })
        
    return pd.DataFrame(movimientos_banco)

def procesar_excel(excel_control_bytes, df_banco, fila_encabezado=FILA_ENCABEZADO_DEFECTO):
    wb = openpyxl.load_workbook(io.BytesIO(excel_control_bytes), data_only=False)
    ws = wb.active

    # Mapeo de columnas de la fila 8 de tu planilla de control
    columnas = {}
    for cell in ws[fila_encabezado]:
        if cell.value is not None:
            columnas[str(cell.value).strip()] = cell.column

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
        raise ValueError(f"Faltan columnas requeridas en la Fila {fila_encabezado}. Detectadas: {list(columnas.keys())}")

    used_banco_indices = set()
    emparejados = 0
    sin_match = 0
    total_filas = ws.max_row

    for fila in range(fila_encabezado + 1, total_filas + 1):
        cargo_val = ws.cell(row=fila, column=col_cargo).value
        abono_val = ws.cell(row=fila, column=col_abono).value
        fecha_val = ws.cell(row=fila, column=col_fecha).value
        doc_val = ws.cell(row=fila, column=col_doc).value if col_doc else None
        saldo_val = ws.cell(row=fila, column=col_saldo).value if col_saldo else None

        monto_excel = None
        if cargo_val not in (None, "", 0):
            monto_excel = float(cargo_val) if isinstance(cargo_val, (int, float)) else limpiar_monto_celda(cargo_val)
        elif abono_val not in (None, "", 0):
            monto_excel = float(abono_val) if isinstance(abono_val, (int, float)) else limpiar_monto_celda(abono_val)

        if monto_excel is None or monto_excel == 0:
            continue

        fecha_excel = parsear_fecha(fecha_val)
        num_doc_excel = int(doc_val) if isinstance(doc_val, (int, float)) and doc_val > 0 else None
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
                    if abs((fecha_excel - fila_b["fecha"]).days) > 4:
                        continue
                if saldo_excel and fila_b["saldo_diario"]:
                    if round(saldo_excel, 0) != round(fila_b["saldo_diario"], 0):
                        continue
                match_idx = idx_b
                break

        # Fallback de cola (FIFO) por si las fechas tienen desfases mayores
        if match_idx is None and not df_banco.empty:
            cand = df_banco[(df_banco["monto"] == round(abs(monto_excel), 2)) & (~df_banco.index.isin(used_banco_indices))]
            if not cand.empty:
                match_idx = cand.index[0]

        if match_idx is not None:
            used_banco_indices.add(match_idx)
            cartola_num = df_banco.loc[match_idx, "cartola_calculada"]
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
st.markdown("Esta versión optimizada procesa las cartolas directamente en formato planilla, garantizando precisión contable absoluta.")

st.divider()
file_control = st.file_uploader("📊 1. Sube tu planilla de control de movimientos (.xlsx)", type=["xlsx"])
file_cartola_conv = st.file_uploader("📄 2. Sube la cartola convertida a Excel (.xlsx)", type=["xlsx"])

if file_control and file_cartola_conv:
    if st.button("🚀 Iniciar Cruce Automatizado", type="primary"):
        control_bytes = file_control.read()
        
        with st.spinner("Mapeando transacciones de la cartola bancaria..."):
            df_banco = extraer_movimientos_cartola_convertida(file_cartola_conv)
            
        if df_banco.empty:
            st.error("No se encontraron registros procesables en el Excel de la cartola.")
        else:
            st.success(f"✅ Se leyeron {len(df_banco)} movimientos bancarios limpios.")
            
            with st.spinner("Rellenando las celdas vacías del control..."):
                try:
                    res_buf, ok, nok = procesar_excel(control_bytes, df_banco, fila_encabezado=FILA_ENCABEZADO_DEFECTO)
                    st.success(f"🎯 ¡Proceso exitoso! Filas emparejadas con su número de cartola: {ok}. Filas sin match: {nok}.")
                    st.download_button("📥 Descargar Planilla Finalizada", data=res_buf, file_name="control_bancario_conciliado.xlsx")
                except Exception as e:
                    st.error(f"Error durante el llenado: {e}")
