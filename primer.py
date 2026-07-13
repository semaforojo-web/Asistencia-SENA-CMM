import streamlit as st
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# ⚠️ CONFIGURACIÓN DIRECTA CON ID DE PESTAÑA ESPECÍFICO
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1tHlKlDD5bVuiZTXhUGAJoJyI8P4bvmRrNjKUXIAK-4g/edit?usp=sharing"
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
            # 1. Cargar las credenciales desde los Secrets
            secret_dict = dict(st.secrets["gspread_credentials"])
            
            if "private_key" in secret_dict:
                secret_dict["private_key"] = secret_dict["private_key"].replace("\\n", "\n")
            
            # 2. DEFINIR LOS ALCANCES (SCOPES) EXPLÍCITOS DE GOOGLE DRIVE Y SHEETS
            # Esto obliga a Google a validar los permisos de lectura y escritura completos
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            
            # 3. Crear las credenciales autorizadas con los alcances definidos
            credentials = Credentials.from_service_account_info(secret_dict, scopes=scopes)
            
            # 4. Conectar el cliente de gspread usando el objeto de credenciales explícito
            client = gspread.authorize(credentials)
            
            # 5. Abrir el libro principal "Reporte de Asistencia"
            spreadsheet = client.open_by_url(URL_GOOGLE_SHEETS)
            
            # Seleccionar la hoja usando directamente el GID numérico brindado
            worksheet = spreadsheet.get_worksheet_by_id(SHEET_GID)
            
            if worksheet is None:
                st.error(f"No se encontró la hoja específica con el ID {SHEET_GID} dentro del archivo.")
            else:
                # Descargar todas las filas de forma segura
                all_values = worksheet.get_all_values()
                
                if not all_values:
                    st.error("La hoja de cálculo está vacía.")
                else:
                    # Convertir a DataFrame
                    df = pd.DataFrame(all_values[1:], columns=all_values[0])
                    df.columns = df.columns.str.strip().str.upper()
                    
                    COL_BUSQUEDA = "NUMERO_DOCUMENTO"
                    
                    if COL_BUSQUEDA not in df.columns:
                        st.error(f"Error técnico: No se encontró la columna '{COL_BUSQUEDA}' en el archivo.")
                        st.info(f"Columnas detectadas: {list(df.columns)}")
                    else:
                        # Asegurar las columnas de destino
                        columnas_requeridas = ['FECHA_REGISTRO', 'DOC_CONFIRMADO', 'CORREO_REGISTRO', 'CELULAR_REGISTRO']
                        for col in columnas_requeridas:
                            if col not in df.columns:
                                df[col] = ""
                        
                        coincidencia_encontrada = False
                        documento_limpio = str(documento).strip()
                        
                        # Buscar el aprendiz
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
                        
                        # Guardar cambios reconstruyendo la matriz
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
