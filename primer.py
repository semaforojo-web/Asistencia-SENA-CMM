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
            # 1. Limpiar la URL base
            url_limpia = URL_GOOGLE_SHEETS.strip().replace(" ", "")
            base_url = url_limpia.split("/edit")[0]
            
            # 2. NOMBRE DE LA HOJA CODIFICADO PARA URL (Cambia los espacios por %20 automáticamente)
            nombre_hoja = "Listado de aprendices"
            nombre_hoja_codificado = urllib.parse.quote(nombre_hoja)
            
            # Construir la URL de exportación segura
            csv_url = f"{base_url}/export?format=csv&sheet={nombre_hoja_codificado}"
            
            # 3. Leer la hoja usando Pandas
            df = pd.read_csv(csv_url)
            
            # Asegurar que existan suficientes columnas para T, U, V, W (índices 19 a 22)
            columnas_necesarias = 23
            while len(df.columns) < columnas_necesarias:
                df[f"Columna_{len(df.columns)}"] = ""
            
            col_L_idx = 11  # Columna L (Número de documento)
            col_T_idx = 19  # Columna T
            col_U_idx = 20  # Columna U
            col_V_idx = 21  # Columna V
            col_W_idx = 22  # Columna W
            
            coincidencia_encontrada = False
            documento_limpio = str(documento).strip()
            
            # 4. Buscar coincidencia en la columna L
            for idx, row in df.iterrows():
                val_L = row.iloc[col_L_idx]
                
                if pd.notna(val_L):
                    # Limpieza del .0 decimal
                    val_L_str = str(val_L).split('.')[0].strip()
                    
                    if val_L_str == documento_limpio:
                        df.iat[idx, col_T_idx] = fecha_hora_local
                        df.iat[idx, col_U_idx] = documento_limpio
                        df.iat[idx, col_V_idx] = correo
                        df.iat[idx, col_W_idx] = celular
                        coincidencia_encontrada = True
            
            if coincidencia_encontrada:
                # 5. Guardar los cambios en Google Sheets
                from streamlit_gsheets import GSheetsConnection
                conn = st.connection("gsheets", type=GSheetsConnection)
                
                # Inyectamos los datos directo al update con el nombre de hoja original
                conn.update(spreadsheet=url_limpia, sheet=nombre_hoja, data=df)
                st.success(f"¡Asistencia registrada con éxito en la nube para el documento {documento}!")
            else:
                st.warning(f"El número de documento '{documento}' no se encontró en la base de datos.")
                
        except Exception as e:
            st.error(f"Error en la operación: {e}")
