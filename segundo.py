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
    else:
        # Llamamos a la función de sincronización directa con GitHub
        exito = guardar_en_github_desde_formulario(documento, correo, celular)
        
        if exito:
            st.success("¡Registro guardado y sincronizado en GitHub!")
        else:
            st.warning("Documento no encontrado o error en la sincronización.")

            
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
def guardar_en_github_desde_formulario(documento, correo, celular):
    """Actualiza el Excel en GitHub con los datos del formulario."""
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("⚠️ Falta el GITHUB_TOKEN en los Secrets.")
        return False

    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo("semaforojo-web/Asistencia-SENA-CMM") # Ajusta según tu repo
        
        # 1. Obtener contenido actual
        contents = repo.get_contents("Reporte de Asistencia.xlsx", ref="main")
        archivo_bytes = io.BytesIO(contents.decoded_content)
        
        # 2. Modificar con openpyxl
        wb = openpyxl.load_workbook(archivo_bytes)
        sheet = wb["Listado de aprendices"]
        
        coincidencia = False
        for row in range(2, sheet.max_row + 1):
            val_L = sheet.cell(row=row, column=12).value
            if val_L and str(val_L).split('.')[0].strip() == documento:
                sheet.cell(row=row, column=20, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                sheet.cell(row=row, column=21, value=documento)
                sheet.cell(row=row, column=22, value=correo)
                sheet.cell(row=row, column=23, value=celular)
                coincidencia = True
                break
        
        if not coincidencia:
            return False

        # 3. Guardar y subir a GitHub
        output = io.BytesIO()
        wb.save(output)
        repo.update_file(
            path="Reporte de Asistencia.xlsx",
            message="🤖 Actualización de datos desde formulario",
            content=output.getvalue(),
            sha=contents.sha,
            branch="main"
        )
        return True
    except Exception as e:
        st.error(f"Error al sincronizar con GitHub: {e}")
        return False
        
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
            
