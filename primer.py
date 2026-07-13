import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import pandas as pd

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
            # 2. Inicializar la conexión con tus Secrets de la Service Account
            conn = st.connection("gsheets", type=GSheetsConnection)
            
            # 3. Acceder al cliente nativo de la API (gspread) para evitar la construcción de URLs propensas a errores
            client = conn._client
            
            # Abrir el libro y la pestaña por su nombre exacto usando la API oficial
            spreadsheet = client.open_by_url(URL_GOOGLE_SHEETS)
            worksheet = spreadsheet.worksheet(SHEET_NAME)
            
            # 4. Descargar todas las filas como una lista de listas de forma segura
            all_values = worksheet.get_all_values()
            
            if not all_values:
                st.error("El archivo de Google Sheets está completamente vacío.")
            else:
                # Convertir a DataFrame. La primera fila se toma como encabezado real
                df = pd.DataFrame(all_values[1:], columns=all_values[0])
                
                # Normalizar nombres de columnas para la búsqueda (quitar espacios y pasar a mayúsculas)
                df.columns = df.columns.str.strip().str.upper()
                
                COL_BUSQUEDA = "NUMERO_DOCUMENTO"
                
                if COL_BUSQUEDA not in df.columns:
                    st.error(f"Error técnico: No se encontró la columna '{COL_BUSQUEDA}' en el archivo.")
                    st.info(f"Columnas detectadas en tu archivo: {list(df.columns)}")
                else:
                    # 5. Asegurar las columnas de destino en mayúsculas si no existen en el DataFrame
                    columnas_requeridas = ['FECHA_REGISTRO', 'DOC_CONFIRMADO', 'CORREO_REGISTRO', 'CELULAR_REGISTRO']
                    for col in columnas_requeridas:
                        if col not in df.columns:
                            df[col] = ""
                    
                    coincidencia_encontrada = False
                    documento_limpio = str(documento).strip()
                    
                    # 6. Buscar el aprendiz recorriendo las filas
                    for idx, row in df.iterrows():
                        val_documento = row[COL_BUSQUEDA]
                        
                        if val_documento:
                            # Limpiar posibles formatos numéricos (.0)
                            val_doc_str = str(val_documento).split('.')[0].strip()
                            
                            if val_doc_str == documento_limpio:
                                df.at[idx, 'FECHA_REGISTRO'] = fecha_hora_local
                                df.at[idx, 'DOC_CONFIRMADO'] = documento_limpio
                                df.at[idx, 'CORREO_REGISTRO'] = correo
                                df.at[idx, 'CELULAR_REGISTRO'] = celular
                                coincidencia_encontrada = True
                    
                    # 7. Si se actualizó el DataFrame, lo guardamos en la hoja de cálculo
                    if coincidencia_encontrada:
                        # Reconstruir la matriz completa incluyendo los encabezados originales corregidos
                        nuevos_encabezados = list(df.columns)
                        nuevos_datos = [nuevos_encabezados] + df.values.tolist()
                        
                        # Limpiar la hoja vieja y escribir la matriz nueva de un solo golpe (Operación CRUD Atómica)
                        worksheet.clear()
                        worksheet.update('A1', nuevos_datos)
                        
                        st.success(f"¡Registro guardado exitosamente en Google Sheets para el documento {documento}!")
                    else:
                        st.warning(f"El número de documento '{documento}' no se encontró en la lista de aprendices.")
                        
        except Exception as e:
            st.error(f"Ocurrió un error técnico al procesar el archivo en la nube: {e}")
