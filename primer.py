import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# ⚠️ ENLACE DIRECTO A TU GOOGLE SHEETS
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
            # Conectar con Google Sheets utilizando la extensión de Streamlit
            conn = st.connection("gsheets", type=GSheetsConnection)
            
            # Leer los datos de la nube en tiempo real (ttl=0 obliga a traer lo último guardado)
            # Usamos header=None para tratar la hoja como una cuadrícula pura por índices de posición
            df = conn.read(spreadsheet=URL_GOOGLE_SHEETS, sheet=SHEET_NAME, ttl=0, header=None)
            
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
            
            # Recorrer las filas de la cuadrícula (empezando desde el índice 1 para ignorar encabezados)
            for idx, row in df.iterrows():
                if idx == 0:
                    continue
                    
                val_L = row.iloc[col_L]
                
                if pd.notna(val_L):
                    # Convertimos a string y dividimos en el punto para remover el ".0" de formato flotante
                    val_L_str = str(val_L).split('.')[0].strip()
                    documento_limpio = str(documento).strip()
                    
                    # Comparación idéntica a la que te funcionó
                    if val_L_str == documento_limpio:
                        df.iat[idx, col_T] = fecha_hora_local
                        df.iat[idx, col_U] = documento_limpio
                        df.iat[idx, col_V] = correo
                        df.iat[idx, col_W] = celular
                        coincidencia_encontrada = True
            
            # 3. Guardar cambios directamente en la nube o reportar error
            if coincidencia_encontrada:
                # Enviamos el dataframe modificado directo a Google Sheets
                conn.update(spreadsheet=URL_GOOGLE_SHEETS, sheet=SHEET_NAME, data=df, headers=False)
                st.success(f"¡Registro guardado exitosamente en Google Sheets para el documento {documento}!")
            else:
                st.warning(f"El número de documento '{documento}' no se encontró en la lista de aprendices.")
                
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {e}")
