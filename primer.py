import streamlit as st
from datetime import datetime
import openpyxl
import os

# Configuración de la página web
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

EXCEL_FILE = "Reporte de Asistencia.xlsx"
SHEET_NAME = "Listado de aprendices"

# --- Componentes del Formulario en la Web ---
with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    
    # Botón de envío dentro del formulario
    enviado = st.form_submit_button("Guardar Registro")

# --- Lógica al presionar el botón ---
if enviado:
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
    elif not os.path.exists(EXCEL_FILE):
        st.error(f"Error: No se encontró el archivo '{EXCEL_FILE}' en el servidor.")
    else:
        # Obtener fecha y hora local del servidor
        fecha_hora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if SHEET_NAME not in wb.sheetnames:
                st.error(f"No se encontró la hoja '{SHEET_NAME}' en el archivo.")
                wb.close()
            else:
                sheet = wb[SHEET_NAME]
                
                # Columnas: L=12, T=20, U=21, V=22, W=23
                col_L = 12
                col_T = 20
                col_U = 21
                col_V = 22
                col_W = 23
                
                coincidencia_encontrada = False
                
                # Buscar en la columna L
                for row in range(2, sheet.max_row + 1):
                    val_L = sheet.cell(row=row, column=col_L).value
                    
                    if val_L is not None and str(val_L).strip() == str(documento):
                        sheet.cell(row=row, column=col_T, value=fecha_hora_local)
                        sheet.cell(row=row, column=col_U, value=str(documento))
                        sheet.cell(row=row, column=col_V, value=correo)
                        sheet.cell(row=row, column=col_W, value=celular)
                        coincidencia_encontrada = True
                
                if coincidencia_encontrada:
                    wb.save(EXCEL_FILE)
                    st.success(f"¡Registro guardado exitosamente para el documento {documento}!")
                else:
                    st.warning(f"El número de documento '{documento}' no se encontró en el listado.")
                
                wb.close()
                
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el Excel: {e}")
            
