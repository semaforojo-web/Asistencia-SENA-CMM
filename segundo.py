import streamlit as st
from datetime import datetime
import openpyxl
import os

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# Configuración del archivo Excel y la hoja destino
EXCEL_FILE = "Reporte de Asistencia.xlsx"
SHEET_NAME = "Listado de aprendices"

# --- Estructura del Formulario en la Web ---
with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    
    # Botón de envío
    enviado = st.form_submit_button("Guardar Registro")

# --- Lógica de procesamiento al presionar el botón ---
if enviado:
    # 1. Validar que no existan campos vacíos
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
        
    # 2. Validar que el archivo Excel esté en el servidor
    elif not os.path.exists(EXCEL_FILE):
        st.error(f"Error: No se encontró el archivo '{EXCEL_FILE}' en el repositorio.")
        
    else:
        # Generar fecha y hora local del sistema
        fecha_hora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            # Cargar el archivo Excel manteniendo estilos previos
            wb = openpyxl.load_workbook(EXCEL_FILE)
            
            if SHEET_NAME not in wb.sheetnames:
                st.error(f"No se encontró la hoja '{SHEET_NAME}' dentro del archivo de Excel.")
                wb.close()
            else:
                sheet = wb[SHEET_NAME]
                
                # Mapeo exacto de las columnas
                col_L = 12
                col_T = 20
                col_U = 21
                col_V = 22
                col_W = 23
                
                coincidencia_encontrada = False
                
                # Recorrer las filas del Excel
                for row in range(2, sheet.max_row + 1):
                    val_L = sheet.cell(row=row, column=col_L).value
                    
                    if val_L is not None:
                        val_L_str = str(val_L).split('.')[0].strip()
                        documento_limpio = str(documento).strip()
                        
                        # Comparación limpia
                        if val_L_str == documento_limpio:
                            sheet.cell(row=row, column=col_T, value=fecha_hora_local)
                            sheet.cell(row=row, column=col_U, value=documento_limpio)
                            sheet.cell(row=row, column=col_V, value=correo)
                            sheet.cell(row=row, column=col_W, value=celular)
                            coincidencia_encontrada = True
                
                # 3. Guardar cambios o reportar error
                if coincidencia_encontrada:
                    wb.save(EXCEL_FILE)
                    st.success(f"¡Registro guardado exitosamente para el documento {documento}!")
                else:
                    st.warning(f"El número de documento '{documento}' no se encontró en la lista.")
                
                wb.close()
                
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo Excel: {e}")
          
