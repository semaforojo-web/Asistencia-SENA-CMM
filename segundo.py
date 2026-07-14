import streamlit as st
from datetime import datetime
import openpyxl
import io
from github import Github

st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
EXCEL_FILE = "Reporte de Asistencia.xlsx"
SHEET_NAME = "Listado de aprendices"

def procesar_y_guardar_github(documento, correo, celular):
    """Lógica corregida para manejar bytes y actualizar en GitHub."""
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("⚠️ Falta el GITHUB_TOKEN en los Secrets.")
        return False, None

    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo("semaforojo-web/Asistencia-SENA-CMM")
        
        # Obtener contenido como bytes directamente[span_1](start_span)[span_1](end_span)
        contents = repo.get_contents(EXCEL_FILE, ref="main")
        archivo_bytes = io.BytesIO(contents.content)
        
        wb = openpyxl.load_workbook(archivo_bytes)
        sheet = wb[SHEET_NAME]
        coincidencia = False
        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Lógica original de detección de documento[span_2](start_span)[span_2](end_span)
        for row in range(2, sheet.max_row + 1):
            val_L = sheet.cell(row=row, column=12).value
            if val_L and str(val_L).split('.')[0].strip() == documento:
                sheet.cell(row=row, column=20, value=fecha_hora)
                sheet.cell(row=row, column=21, value=documento)
                sheet.cell(row=row, column=22, value=correo)
                sheet.cell(row=row, column=23, value=celular)
                coincidencia = True
        
        if coincidencia:
            # Guardar en memoria como bytes[span_3](start_span)[span_3](end_span)[span_4](start_span)[span_4](end_span)
            output = io.BytesIO()
            wb.save(output)
            contenido_final = output.getvalue() 
            
            # Subir a GitHub pasando bytes, no un string[span_5](start_span)[span_5](end_span)
            repo.update_file(
                path=EXCEL_FILE,
                message="🤖 Actualización desde formulario",
                content=contenido_final,
                sha=contents.sha,
                branch="main"
            )
            return True, contenido_final
        return False, None
    except Exception as e:
        st.error(f"Error al sincronizar con GitHub: {e}")
        return False, None

with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    enviado = st.form_submit_button("Guardar Registro")

if enviado:
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
    else:
        exito, archivo_bytes = procesar_y_guardar_github(documento, correo, celular)
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
            
