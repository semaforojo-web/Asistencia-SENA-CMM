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
        
        # 1. Obtener contenido crudo (bytes)[span_4](start_span)[span_4](end_span)
        contents = repo.get_contents(EXCEL_FILE, ref="main")
        archivo_bytes = io.BytesIO(contents.content)
        
        # 2. Procesar el Excel en memoria
        wb = openpyxl.load_workbook(archivo_bytes)
        sheet = wb[SHEET_NAME]
        
        coincidencia = False
        documento_str = str(documento).strip()
        
        # 3. Buscar coincidencia (manejo estricto de tipos de datos)[span_5](start_span)[span_5](end_span)
        for row in range(2, sheet.max_row + 1):
            val_L = sheet.cell(row=row, column=12).value
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

        # 4. Guardar en memoria y preparar bytes para GitHub[span_6](start_span)[span_6](end_span)
        output = io.BytesIO()
        wb.save(output)
        contenido_bytes = output.getvalue() 
        
        # 5. Subir a GitHub[span_7](start_span)[span_7](end_span)
        repo.update_file(
            path=EXCEL_FILE,
            message="🤖 Actualización de datos desde formulario",
            content=contenido_bytes,
            sha=contents.sha,
            branch="main"
        )
        return True, contenido_bytes
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
            
