import streamlit as st
from datetime import datetime
import openpyxl
import os
import io
from github import Github
import pandas as pd


st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
EXCEL_FILE = "Reporte de Asistencia.xlsx"
SHEET_NAME = "Listado de aprendices"

with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    enviado = st.form_submit_button("Guardar Registro")

if enviado:
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
    elif not os.path.exists(EXCEL_FILE):
        st.error(f"Error: No se encontró el archivo '{EXCEL_FILE}'.")
    else:
        fecha_hora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            sheet = wb[SHEET_NAME]
            coincidencia = False
            
            for row in range(2, sheet.max_row + 1):
                val_L = sheet.cell(row=row, column=12).value
                if val_L and str(val_L).split('.')[0].strip() == documento:
                    sheet.cell(row=row, column=20, value=fecha_hora_local)
                    sheet.cell(row=row, column=21, value=documento)
                    sheet.cell(row=row, column=22, value=correo)
                    sheet.cell(row=row, column=23, value=celular)
                    coincidencia = True
            
            if coincidencia:
                wb.save(EXCEL_FILE)
                st.success("¡Registro guardado!")
                
                # Botón de descarga para el archivo modificado
                with open(EXCEL_FILE, "rb") as file:
                    st.download_button(
                        label="Descargar Reporte Actualizado",
                        data=file,
                        file_name="Reporte_Asistencia_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.warning("Documento no encontrado.")
            wb.close()
        except Exception as e:
            st.error(f"Error técnico: {e}")
            
