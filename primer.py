import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# ⚠️ ENLACES DIRECTOS
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1tHlKlDD5bVuiZTXhUGAJoJyI8P4bvmRrNjKUXIAK-4g/edit?usp=sharing"
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
        
    else:
        # Generar fecha y hora local del sistema
        fecha_hora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            # 2. Inicializar la conexión usando las credenciales seguras de "Secrets"
            conn = st.connection("gsheets", type=GSheetsConnection)
            
            # 3. Leer los datos de forma segura autenticada (usando los encabezados reales de texto)
            df = conn.read(spreadsheet=URL_GOOGLE_SHEETS, worksheet=SHEET_NAME, ttl=0)
            
            # Convertir todas las columnas de texto a mayúsculas y limpiar espacios para evitar discrepancias
            df.columns = df.columns.str.strip().str.upper()
            
            # Validar que exista la columna de documento (gracias a la Service Account se lee de forma nativa)
            COL_BUSQUEDA = "NUMERO_DOCUMENTO"
            if COL_BUSQUEDA not in df.columns:
                st.error(f"Error técnico: No se encontró la columna '{COL_BUSQUEDA}' en el archivo.")
            else:
                # 4. Asegurar la existencia de las columnas de destino T, U, V, W si no existen
                # Si trabajas por nombres en Sheets autenticado, puedes mapearlas directamente o asegurar el ancho de la matriz:
                columnas_requeridas = ['FECHA_REGISTRO', 'DOC_CONFIRMADO', 'CORREO_REGISTRO', 'CELULAR_REGISTRO']
                for col in columnas_requeridas:
                    if col not in df.columns:
                        df[col] = ""
                
                coincidencia_encontrada = False
                documento_limpio = str(documento).strip()
                
                # 5. Recorrer las filas buscando la coincidencia utilizando los nombres reales de las columnas
                for idx, row in df.iterrows():
                    val_documento = row[COL_BUSQUEDA]
                    
                    if pd.notna(val_documento):
                        # Quitar el decimal .0 si Excel o Sheets lo interpretó numéricamente
                        val_doc_str = str(val_documento).split('.')[0].strip()
                        
                        if val_doc_str == documento_limpio:
                            df.at[idx, 'FECHA_REGISTRO'] = fecha_hora_local
                            df.iat[idx, df.columns.get_loc('DOC_CONFIRMADO')] = documento_limpio
                            df.iat[idx, df.columns.get_loc('CORREO_REGISTRO')] = correo
                            df.iat[idx, df.columns.get_loc('CELULAR_REGISTRO')] = celular
                            coincidencia_encontrada = True
                
                # 6. Guardar los datos modificados directamente en la nube
                if coincidencia_encontrada:
                    conn.update(spreadsheet=URL_GOOGLE_SHEETS, worksheet=SHEET_NAME, data=df)
                    st.success(f"¡Registro guardado exitosamente en Google Sheets para el documento {documento}!")
                else:
                    st.warning(f"El número de documento '{documento}' no se encontró en la base de datos de aprendices.")
                
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {e}")
