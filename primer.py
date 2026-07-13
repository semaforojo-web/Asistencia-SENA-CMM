import streamlit as st
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# 🎯 CONFIGURACIÓN MEDIANTE ID DE DOCUMENTO (MÁS SEGURO)
# Extraído directamente de tu enlace de Google Sheets
SPREADSHEET_KEY = "1tHlKlDD5bVuiZTXhUGAJoJyI8P4bvmRrNjKUXIAK-4g"
SHEET_GID = 601595677  # ID de pestaña específico: Listado de Aprendices

# --- Estructura del Formulario en la Web ---
with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    
    # Botón de envío
    enviado = st.form_submit_button("Guardar Registro")

# --- Lógica de procesamiento al presionar el botón ---
if enviado:
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
        
    else:
        fecha_hora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            # Limpieza explícita de cualquier caché previa
            st.cache_resource.clear()
            
            # Cargar las credenciales desde los Secrets
            secret_dict = dict(st.secrets["gspread_credentials"])
            
            if "private_key" in secret_dict:
                secret_dict["private_key"] = secret_dict["private_key"].replace("\\n", "\n")
            
            # Definir alcances explícitos
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            
            # Autorizar el cliente
            credentials = Credentials.from_service_account_info(secret_dict, scopes=scopes)
            client = gspread.authorize(credentials)
            
            # 🎯 CAMBIO AQUÍ: Abrimos directamente usando la KEY del documento en vez de la URL
            spreadsheet = client.open_by_key(SPREADSHEET_KEY)
            worksheet = spreadsheet.get_worksheet_by_id(SHEET_GID)
            
            if worksheet is None:
                st.error(f"No se encontró la hoja específica con el ID {SHEET_GID} dentro del archivo.")
            else:
                all_values = worksheet.get_all_values()
                
                if not all_values:
                    st.error("La hoja de cálculo está vacía.")
                else:
                    df = pd.DataFrame(all_values[1:], columns=all_values[0])
                    df.columns = df.columns.str.strip().str.upper()
                    
                    COL_BUSQUEDA = "NUMERO_DOCUMENTO"
                    
                    if COL_BUSQUEDA not in df.columns:
                        st.error(f"Error técnico: No se encontró la columna '{COL_BUSQUEDA}' en el archivo.")
                    else:
                        columnas_requeridas = ['FECHA_REGISTRO', 'DOC_CONFIRMADO', 'CORREO_REGISTRO', 'CELULAR_REGISTRO']
                        for col in columnas_requeridas:
                            if col not in df.columns:
                                df[col] = ""
                        
                        coincidencia_encontrada = False
                        documento_limpio = str(documento).strip()
                        
                        for idx, row in df.iterrows():
                            val_documento = row[COL_BUSQUEDA]
                            if val_documento:
                                val_doc_str = str(val_documento).split('.')[0].strip()
                                if val_doc_str == documento_limpio:
                                    df.at[idx, 'FECHA_REGISTRO'] = fecha_hora_local
                                    df.at[idx, 'DOC_CONFIRMADO'] = documento_limpio
                                    df.at[idx, 'CORREO_REGISTRO'] = correo
                                    df.at[idx, 'CELULAR_REGISTRO'] = celular
                                    coincidencia_encontrada = True
                        
                        if coincidencia_encontrada:
                            nuevos_encabezados = list(df.columns)
                            nuevos_datos = [nuevos_encabezados] + df.values.tolist()
                            
                            worksheet.clear()
                            worksheet.update(range_name='A1', values=nuevos_datos)
                            
                            st.success(f"¡Registro guardado exitosamente para el documento {documento}!")
                        else:
                            st.warning(f"El número de documento '{documento}' no se encontró en la lista de aprendices.")
                            
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {str(e) if str(e) else type(e).__name__}")
