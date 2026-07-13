import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Configuración de la interfaz web
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización (Google Sheets)")
st.write("Ingrese los datos solicitados para registrar su asistencia en tiempo real.")

# ⚠️ REEMPLAZA ESTA URL CON EL ENLACE QUE COPIASTE DE TU GOOGLE SHEETS
# Asegúrate de mantener las comillas.
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
            # 1. Conectar con el Google Sheets utilizando st.connection
            conn = st.connection("gsheets", type=GSheetsConnection)
            
            # 2. Leer la hoja exacta "Listado de aprendices"
            # Ponemos ttl=0 para que siempre traiga los datos más frescos de la nube
            df = conn.read(spreadsheet=URL_GOOGLE_SHEETS, sheet="Listado de aprendices", ttl=0)
            
            # Nos aseguramos de rellenar las columnas si la hoja original es más corta
            # Columnas requeridas en base a letras de Excel: L es columna 11 (0-indexed)
            # T, U, V, W corresponden a los índices 19, 20, 21, 22
            columnas_necesarias = 23
            while len(df.columns) < columnas_necesarias:
                df[f"Columna_{len(df.columns)}"] = ""
                
            col_L_idx = 11
            col_T_idx = 19
            col_U_idx = 20
            col_V_idx = 21
            col_W_idx = 22
            
            coincidencia_encontrada = False
            documento_limpio = str(documento).strip()
            
            # 3. Buscar coincidencia en la columna L (index 11)
            for idx, row in df.iterrows():
                val_L = row.iloc[col_L_idx]
                
                if pd.notna(val_L):
                    # Limpieza estándar para evitar problemas de formato (.0 decimales)
                    val_L_str = str(val_L).split('.')[0].strip()
                    
                    if val_L_str == documento_limpio:
                        # Convertir la fila en lista modificable si pandas la bloquea
                        df.iat[idx, col_T_idx] = fecha_hora_local
                        df.iat[idx, col_U_idx] = documento_limpio
                        df.iat[idx, col_V_idx] = correo
                        df.iat[idx, col_W_idx] = celular
                        coincidencia_encontrada = True
            
            # 4. Guardar los cambios directamente en la nube si fue encontrado
            if coincidencia_encontrada:
                # Actualizar la hoja de cálculo en la nube
                conn.update(spreadsheet=URL_GOOGLE_SHEETS, sheet="Listado de aprendices", data=df)
                st.success(f"¡Asistencia registrada con éxito en la nube para el documento {documento}!")
            else:
                st.warning(f"El número de documento '{documento}' no se encontró en la base de datos.")
                
        except Exception as e:
            st.error(f"Error al conectar con Google Sheets: {e}")
