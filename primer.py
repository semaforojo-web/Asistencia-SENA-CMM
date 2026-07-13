import streamlit as st
from datetime import datetime
import pandas as pd
import urllib.parse

# Configuración de la interfaz web
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia en tiempo real.")

# URL de tu Google Sheets
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1tHlKlDD5bVuiZTXhUGAJoJyI8P4bvmRrNjKUXIAK-4g/edit?usp=sharing"

# --- Estructura del Formulario en la Web ---
with st.form("formulario_asistencia", clear_on_submit=True):
    documento = st.text_input("Número de Documento:").strip()
    correo = st.text_input("Correo Electrónico:").strip()
    celular = st.text_input("Número de Celular:").strip()
    
    enviado = st.form_submit_button("Guardar Registro")

# --- Lógica de procesamiento ---
if enviado:
    if not documento or not correo or not celular:
        st.error("Todos los campos son obligatorios.")
    else:
        fecha_hora_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            # 1. Limpiar y estructurar la URL de descarga
            url_limpia = URL_GOOGLE_SHEETS.strip().replace(" ", "")
            base_url = url_limpia.split("/edit")[0]
            
            nombre_hoja = "Listado de aprendices"
            nombre_hoja_codificado = urllib.parse.quote(nombre_hoja)
            csv_url = f"{base_url}/export?format=csv&sheet={nombre_hoja_codificado}"
            
            # 2. Leer la tabla como una cuadrícula pura sin nombres de columna (header=None)
            # Esto hace que la fila 1 (los encabezados) sea simplemente la fila con índice 0.
            df = pd.read_csv(csv_url, header=None)
            
            # 3. Asegurar que la estructura cuente con espacio suficiente hasta la columna W (23 columnas)
            while len(df.columns) < 23:
                df[len(df.columns)] = ""
            
            # MAPEO ASOCIADO DIRECTO A LAS LETRAS DE EXCEL:
            # Columna L = Índice 11
            # Columna T = Índice 19
            # Columna U = Índice 20
            # Columna V = Índice 21
            # Columna W = Índice 22
            col_L_idx = 11
            col_T_idx = 19
            col_U_idx = 20
            col_V_idx = 21
            col_W_idx = 22
            
            coincidencia_encontrada = False
            documento_limpio = str(documento).strip()
            
            # 4. Buscar coincidencia recorriendo cada una de las filas
            for idx, row in df.iterrows():
                # Omitimos la primera fila (índice 0) ya que contiene las etiquetas originales escritas
                if idx == 0:
                    continue
                    
                val_L = row.iloc[col_L_idx]
                
                if pd.notna(val_L):
                    # Limpieza estándar para quitar el ".0" decimal automático de las celdas numéricas
                    val_L_str = str(val_L).split('.')[0].strip()
                    
                    if val_L_str == documento_limpio:
                        df.iat[idx, col_T_idx] = fecha_hora_local
                        df.iat[idx, col_U_idx] = documento_limpio
                        df.iat[idx, col_V_idx] = correo
                        df.iat[idx, col_W_idx] = celular
                        coincidencia_encontrada = True
            
            # 5. Guardar los cambios si fue hallado con éxito
            if coincidencia_encontrada:
                from streamlit_gsheets import GSheetsConnection
                conn = st.connection("gsheets", type=GSheetsConnection)
                
                # Al actualizar, le indicamos headers=False para que mantenga la fila de títulos intacta
                conn.update(spreadsheet=url_limpia, sheet=nombre_hoja, data=df, headers=False)
                st.success(f"¡Asistencia registrada con éxito en la nube para el documento {documento}!")
            else:
                st.warning(f"El número de documento '{documento}' no se encuentra registrado en la columna L de la base de datos.")
                
        except Exception as e:
            st.error(f"Error en la operación: {e}")
