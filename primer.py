import streamlit as st
from datetime import datetime
import pandas as pd
import gspread

# Configuración de la interfaz web de la página
st.set_page_config(page_title="Registro de Aprendices - SENA", page_icon="📝")

st.title("Formulario de Asistencia / Actualización")
st.write("Ingrese los datos solicitados para registrar su asistencia.")

# ⚠️ CONFIGURACIÓN DIRECTA
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
            # 2. AUTENTICACIÓN OFICIAL DIRECTA
            # Leemos las credenciales directamente como un diccionario nativo de Streamlit Secrets
            secret_dict = dict(st.secrets["gspread_credentials"])
            
            # gspread necesita que la llave privada procese correctamente los saltos de línea (\n)
            if "private_key" in secret_dict:
                secret_dict["private_key"] = secret_dict["private_key"].replace("\\n", "\n")
            
            # Conectamos usando gspread
            client = gspread.service_account_from_dict(secret_dict)
            
            # 3. Abrir el libro y la pestaña usando la API oficial
            spreadsheet = client.open_by_url(URL_GOOGLE_SHEETS)
            worksheet = spreadsheet.worksheet(SHEET_NAME)
            
            # 4. Descargar todas las filas de forma segura
            all_values = worksheet.get_all_values()
            
            if not all_values:
                st.error("El archivo de Google Sheets está completamente vacío.")
            else:
                # Convertir a DataFrame
                df = pd.DataFrame(all_values[1:], columns=all_values[0])
                df.columns = df.columns.str.strip().str.upper()
                
                COL_BUSQUEDA = "NUMERO_DOCUMENTO"
                
                if COL_BUSQUEDA not in df.columns:
                    st.error(f"Error técnico: No se encontró la columna '{COL_BUSQUEDA}' en el archivo.")
                else:
                    # 5. Asegurar las columnas de destino
                    columnas_requeridas = ['FECHA_REGISTRO', 'DOC_CONFIRMADO', 'CORREO_REGISTRO', 'CELULAR_REGISTRO']
                    for col in columnas_requeridas:
                        if col not in df.columns:
                            df[col] = ""
                    
                    coincidencia_encontrada = False
                    documento_limpio = str(documento).strip()
                    
                    # 6. Buscar el aprendiz
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
                    
                    # 7. Guardar en la hoja de cálculo
                    if coincidencia_encontrada:
                        nuevos_encabezados = list(df.columns)
                        nuevos_datos = [nuevos_encabezados] + df.values.tolist()
                        
                        worksheet.clear()
                        worksheet.update(range_name='A1', values=nuevos_datos)
                        
                        st.success(f"¡Registro guardado exitosamente en Google Sheets para el documento {documento}!")
                    else:
                        st.warning(f"El número de documento '{documento}' no se encontró en la lista de aprendices.")
                        
        except Exception as e:
            # Si ocurre un error, ahora sí nos forzará a ver un mensaje descriptivo en texto plano
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {str(e) if str(e) else type(e).__name__}")
