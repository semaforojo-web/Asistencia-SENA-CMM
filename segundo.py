import streamlit as st
from datetime import datetime
import openpyxl
import io
import base64  # Necesario para la codificación correcta
from github import Github

# Configuración de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
EXCEL_FILE = "Reporte de Asistencia.xlsx"
SHEET_NAME = "Listado de aprendices"

def procesar_y_guardar_github(documento, correo, celular):
    """Lógica corregida: utiliza base64 para evitar errores de tipo de datos."""
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("⚠️ Falta el GITHUB_TOKEN en los Secrets.")
        return False, None

    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo("semaforojo-web/Asistencia-SENA-CMM")
        
        # 1. Obtener contenido como bytes[span_2](start_span)[span_2](end_span)
        contents = repo.get_contents(EXCEL_FILE, ref="main")
        archivo_bytes = io.BytesIO(contents.content)
        
        wb = openpyxl.load_workbook(archivo_bytes)
        sheet = wb[SHEET_NAME]
        coincidencia = False
        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 2. Lógica de detección de documento original[span_3](start_span)[span_3](end_span)
        for row in range(2, sheet.max_row + 1):
            val_L = sheet.cell(row=row, column=12).value
            if val_L and str(val_L).split('.')[0].strip() == documento:
                sheet.cell(row=row, column=20, value=fecha_hora)
                sheet.cell(row=row, column=21, value=documento)
                sheet.cell(row=row, column=22, value=correo)
                sheet.cell(row=row, column=23, value=celular)
                coincidencia = True
        
        if coincidencia:
            # 3. Guardar en memoria y preparar bytes[span_4](start_span)[span_4](end_span)[span_5](start_span)[span_5](end_span)
            output = io.BytesIO()
            wb.save(output)
            contenido_binario = output.getvalue()
            
            # 4. Codificar a base64 para evitar errores de tipo de datos en la API[span_6](start_span)[span_6](end_span)
            contenido_b64 = base64.b64encode(contenido_binario).decode("utf-8")
            
            # 5. Subir a GitHub[span_7](start_span)[span_7](end_span)
            repo.update_file(
                path=EXCEL_FILE,
                message="🤖 Actualización desde formulario",
                content=contenido_b64,
                sha=contents.sha,
                branch="main"
            )
            return True, contenido_binario
        return False, None
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
            
