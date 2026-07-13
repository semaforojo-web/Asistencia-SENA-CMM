import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import urllib.parse

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# URL de tu Google Sheets y nombre de la hoja
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
            # 2. Limpiar la URL base y codificar el nombre de la hoja para internet
            url_limpia = URL_GOOGLE_SHEETS.strip().replace(" ", "")
            base_url = url_limpia.split("/edit")[0]
            nombre_hoja_codificado = urllib.parse.quote(SHEET_NAME)
            
            # Construir la URL de exportación directa en formato CSV
            csv_url = f"{base_url}/export?format=csv&sheet={nombre_hoja_codificado}"
            
            # 3. Leer la hoja de Google Sheets de forma nativa e infalible (evitando el bug de conn.read)
            # Usamos header=None para procesar el Excel como una cuadrícula pura por índices de posición
            df = pd.read_csv(csv_url, header=None)
            
            # Asegurar la existencia de las columnas de destino T, U, V, W (23 columnas en total)
            while len(df.columns) < 23:
                df[len(df.columns)] = ""
            
            # Mapeo exacto de las columnas solicitadas por índice de posición (0-indexed):
            # Columna L (Número de documento) = Índice 11
            # Columnas de destino: T = 19, U = 20, V = 21, W = 22
            col_L = 11
            col_T = 19
            col_U = 20
            col_V = 21
            col_W = 22
            
            coincidencia_encontrada = False
            
            # 4. Recorrer las filas de la cuadrícula (empezando desde el índice 1 para ignorar encabezados)
            for idx, row in df.iterrows():
                if idx == 0:
                    continue
                    
                val_L = row.iloc[col_L]
                
                if pd.notna(val_L):
                    # Convertimos a string y dividimos en el punto para remover el ".0" de formato flotante
                    val_L_str = str(val_L).split('.')[0].strip()
                    documento_limpio = str(documento).strip()
                    
                    # Comparación idéntica
                    if val_L_str == documento_limpio:
                        df.iat[idx, col_T] = fecha_hora_local
                        df.iat[idx, col_U] = documento_limpio
                        df.iat[idx, col_V] = correo
                        df.iat[idx, col_W] = celular
                        coincidencia_encontrada = True
            
            # 5. Guardar cambios directamente en la nube si hubo éxito
            if coincidencia_encontrada:
                # Inicializar la conexión solo para empujar los datos actualizados
                conn = st.connection("gsheets", type=GSheetsConnection)
                
                # Enviamos el dataframe modificado directo a Google Sheets usando la URL limpia
                conn.update(spreadsheet=url_limpia, sheet=SHEET_NAME, data=df, headers=False)
                st.success(f"¡Registro guardado exitosamente en Google Sheets para el documento {documento}!")
            else:
                st.warning(f"El número de documento '{documento}' no se encontró en la lista de aprendices.")
                
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {e}")
            
