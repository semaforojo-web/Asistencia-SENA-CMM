import streamlit as st
from datetime import datetime
import openpyxl
import io
from github import Github
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
EXCEL_FILE = "Reporte de Asistencia.xlsx"
SHEET_NAME = "Listado de aprendices"

def guardar_en_github_desde_formulario(documento, correo, celular):
    """Actualiza el Excel en GitHub manejando tipos de datos y errores de codificación."""
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("⚠️ Falta el GITHUB_TOKEN en los Secrets.")
        return False, None

    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo("semaforojo-web/Asistencia-SENA-CMM")
        
        # Obtener contenido crudo (bytes) para evitar error de encoding[span_1](start_span)[span_1](end_span)
        contents = repo.get_contents(EXCEL_FILE, ref="main")
        archivo_bytes = io.BytesIO(contents.content)
        
        wb = openpyxl.load_workbook(archivo_bytes)
        sheet = wb[SHEET_NAME]
        
        coincidencia = False
        documento_str = str(documento).strip()
        
        for row in range(2, sheet.max_row + 1):
            val_L = sheet.cell(row=row, column=12).value
            
            # Limpieza para comparar: convertir a string y eliminar posibles decimales[span_2](start_span)[span_2](end_span)
            val_L_str = str(val_L).split('.')[0].strip()
            
            if val_L_str == documento_str:
                sheet.cell(row=row, column=20, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                sheet.cell(row=row, column=21, value=documento)
                sheet.cell(row=row, column=22, value=correo)
                sheet.cell(row=row, column=23, value=celular)
                coincidencia = True
                break
        
        if not coincidencia:
            return False, None

        # Guardar en memoria y subir a GitHub
        output = io.BytesIO()
        wb.save(output)
        
        repo.update_file(
            path=EXCEL_FILE,
            message="🤖 Actualización de datos desde formulario",
            content=output.getvalue(),
            sha=contents.sha,
            branch="main"
        )
        return True, output.getvalue()
    except Exception as e:
        st.error(f"Error al sincronizar con GitHub: {e}")
        return False, None

# Interfaz del Formulario
with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    enviado = st.form_submit_button("Guardar Registro")

if enviado:
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
    else:
        exito, archivo_bytes = guardar_en_github_desde_formulario(documento, correo, celular)
        
        if exito:
            st.success("¡Registro guardado y sincronizado en GitHub!")
            st.download_button(
                label="Descargar Reporte Actualizado",
                data=archivo_bytes,
                file_name="Reporte_Asistencia_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Documento no encontrado o error en la sincronización.")
            
