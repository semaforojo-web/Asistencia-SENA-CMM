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
            
            # Codificar el nombre de la hoja para internet
            nombre_hoja = "Listado de aprendices"
            nombre_hoja_codificado = urllib.parse.quote(nombre_hoja)
            csv_url = f"{base_url}/export?format=csv&sheet={nombre_hoja_codificado}"
            
            # 2. Leer la hoja usando Pandas
            df = pd.read_csv(csv_url)
            
            # 3. BUSCAR LA COLUMNA POR SU NOMBRE TEXTUAL (Evita errores si se mueve la posición)
            columna_documento = "NUMERO_DOCUMENTO"
            
            if columna_documento not in df.columns:
                st.error(f"No se encontró la columna '{columna_documento}' en tu hoja de Google Sheets. Verifica los encabezados.")
            else:
                # Asegurar la existencia de las columnas de destino T, U, V, W mediante posiciones absolutas en Excel
                # Excel cuenta: A=1, B=2 ... L=12 ... T=20, U=21, V=22, W=23. En Python (0-indexed) son 19, 20, 21, 22.
                while len(df.columns) < 23:
                    df[f"Columna_Nueva_{len(df.columns)+1}"] = ""
                
                # Asignamos los índices absolutos de escritura para asegurar que caigan en T, U, V, W
                col_T_idx = 19
                col_U_idx = 20
                col_V_idx = 21
                col_W_idx = 22
                
                coincidencia_encontrada = False
                documento_limpio = str(documento).strip()
                
                # 4. Recorrer las filas comparando los valores
                for idx, row in df.iterrows():
                    val_L = row[columna_documento]
                    
                    if pd.notna(val_L):
                        # Limpiar formatos numéricos raros (.0 decimales que pone Excel de forma automática)
                        val_L_str = str(val_L).split('.')[0].strip()
                        
                        if val_L_str == documento_limpio:
                            # Escribir los datos en los casilleros de las comunas T, U, V, W respectivamente
                            df.iat[idx, col_T_idx] = fecha_hora_local
                            df.iat[idx, col_U_idx] = documento_limpio
                            df.iat[idx, col_V_idx] = correo
                            df.iat[idx, col_W_idx] = celular
                            coincidencia_encontrada = True
                
                # 5. Guardar los cambios si hubo éxito
                if coincidencia_encontrada:
                    from streamlit_gsheets import GSheetsConnection
                    conn = st.connection("gsheets", type=GSheetsConnection)
                    
                    # Sobrescribir la base de datos en Google Sheets de forma directa
                    conn.update(spreadsheet=url_limpia, sheet=nombre_hoja, data=df)
                    st.success(f"¡Asistencia registrada con éxito en la nube para el documento {documento}!")
                else:
                    st.warning(f"El número de documento '{documento}' no se encuentra registrado en el listado de aprendices.")
                    
        except Exception as e:
            st.error(f"Error en la operación: {e}")
