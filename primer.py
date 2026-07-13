import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1tHlKlDD5bVuiZTXhUGAJoJyI8P4bvmRrNjKUXIAK-4g/edit?usp=sharing"

GID_HOJA = 601595677 

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
            # 2. Limpiar y estructurar la URL de descarga directa usando el número GID libre de espacios
            url_limpia = URL_GOOGLE_SHEETS.strip().replace(" ", "")
            base_url = url_limpia.split("/edit")[0]
            csv_url = f"{base_url}/export?format=csv&gid={GID_HOJA}"
            
            # 3. Descargar la cuadrícula de datos usando el ID numérico exacto
            df = pd.read_csv(csv_url, header=None)
            
            # Asegurar la existencia de las columnas de destino T, U, V, W (23 columnas en total)
            while len(df.columns) < 23:
                df[len(df.columns)] = ""
            
            # Mapeo exacto por posiciones físicas (Índices basados en 0):
            # Columna L (Número de documento) = Índice 11
            # Columnas de destino: T = 19, U = 20, V = 21, W = 22
            col_L = 11
            col_T = 19
            col_U = 20
            col_V = 21
            col_W = 22
            
            coincidencia_encontrada = False
            documento_limpio = str(documento).strip()
            
            # 4. Recorrer las filas buscando la coincidencia
            for idx, row in df.iterrows():
                if idx == 0:
                    continue  # Ignorar los títulos
                    
                val_L = row.iloc[col_L]
                
                if pd.notna(val_L):
                    # Quitar decimales flotantes como el ".0"
                    val_L_str = str(val_L).split('.')[0].strip()
                    
                    if val_L_str == documento_limpio:
                        df.iat[idx, col_T] = fecha_hora_local
                        df.iat[idx, col_U] = documento_limpio
                        df.iat[idx, col_V] = correo
                        df.iat[idx, col_W] = celular
                        coincidencia_encontrada = True
            
            # 5. Guardar los datos modificados directamente en la nube
            if coincidencia_encontrada:
                conn = st.connection("gsheets", type=GSheetsConnection)
                # Actualizamos usando la referencia limpia de red libre de caracteres especiales
                conn.update(spreadsheet=url_limpia, spreadsheet_id=GID_HOJA, data=df, headers=False)
                st.success(f"¡Registro guardado exitosamente en Google Sheets para el documento {documento}!")
            else:
                st.warning(f"El número de documento '{documento}' no se encontró en la columna L de la lista.")
                
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {e}")
