import pandas as pd
import os
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="Control de Asistencia y Evaluación - SENA", layout="wide")

# Nombre del archivo original de Excel
DB_FILE = "Reporte de Asistencia.xlsx"

def cargar_datos():
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None)
            df_procesado = pd.DataFrame()
            
            # Columna E (Ficha) -> Posición 4
            df_procesado["Ficha"] = df.iloc[:, 4].astype(str).str.strip()
            # Columna L (Documento) -> Posición 11
            df_procesado["Documento"] = df.iloc[:, 11].fillna("S/D").astype(str).str.strip()
            # Columnas M a O (Nombres y Apellidos) -> Posiciones 12 a 14
            nombres_completos = df.iloc[:, 12:15].fillna("").astype(str)
            df_procesado["Nombre Completo"] = nombres_completos.agg(' '.join, axis=1).str.replace(r'\s+', ' ', regex=True).str.strip()
            
            # FILTRADO POR COLUMNA P (Posición 15)
            df_procesado["Estado"] = df.iloc[:, 15].fillna("").astype(str).str.upper().str.strip()
            terminos_excluir = "CANCELADO|RETIRO VOLUNTARIO|TRASLADO"
            df_procesado = df_procesado[~df_procesado["Estado"].str.contains(terminos_excluir, regex=True)]
            
            df_procesado = df_procesado[df_procesado["Ficha"].str.contains(r'^\d+$', na=False)]
            df_procesado["Nombre Completo"] = df_procesado["Nombre Completo"].replace("", "Aprendiz sin nombre registrado")
            
            return df_procesado[["Ficha", "Documento", "Nombre Completo"]].reset_index(drop=True)
        except Exception as e:
            st.error(f"Error al filtrar listado de aprendices: {e}")
            
    data = {"Ficha": ["Error"], "Documento": ["0"], "Nombre Completo": ["ARCHIVO EXCEL NO DETECTADO"]}
    return pd.DataFrame(data)

def obtener_parametros_cabezote():
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            instructores = df_cab.iloc[:, 5].dropna().astype(str).str.strip().unique().tolist()
            instructores = [i for i in instructores if i.upper() != "INSTRUCTOR" and i != ""]
            return sorted(instructores)
        except Exception:
            pass
    return ["No Detectado"]

def filtrar_materia_cabezote(ficha, instructor, materia_num):
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[3] = df_cab[3].astype(str).str.strip()
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            
            filtro = (df_cab[6] == str(ficha)) & \
                     (df_cab[5] == str(instructor).upper()) & \
                     (df_cab[3] == str(materia_num))
                     
            resultado = df_cab[filtro]
            
            if not resultado.empty:
                materia_texto = resultado.iloc[0, 10] if resultado.shape[1] > 10 else resultado.iloc[0, 3]
                if pd.notna(materia_texto) and str(materia_texto).strip() != "":
                    return str(materia_texto).strip()
            return f"Materia {materia_num} (Sin descripción detallada en Cabezote)"
        except Exception as e:
            return f"Error leyendo Cabezote: {e}"
    return "Archivo Excel no encontrado"

# --- EJECUCIÓN PRINCIPAL ---
df_aprendices = cargar_datos()
lista_instructores = obtener_parametros_cabezote()

if not os.path.exists("asistencia_guardada.csv"):
    pd.DataFrame(columns=["Fecha", "Ficha", "Instructor", "Materia_Num", "Materia_Nombre", "Documento", "Nombre", "Asistencia"]).to_csv("asistencia_guardada.csv", index=False)

# --- INTERFAZ SIDEBAR ---
st.title("📊 Dashboard de Gestión de Ambientes de Formación - SENA")
st.sidebar.header("⚙️ Filtros de Planificación")

lista_fichas = sorted(df_aprendices["Ficha"].dropna().unique())
ficha_seleccionada = st.sidebar.selectbox("1. Seleccione la Ficha (Columna G):", lista_fichas)
instructor_seleccionado = st.sidebar.selectbox("2. Seleccione el Instructor (Columna F):", lista_instructores)
materia_num_seleccionada = st.sidebar.selectbox("3. Seleccione Materia (Columna D):", ["1", "2", "3"])

materia_detectada = filtrar_materia_cabezote(ficha_seleccionada, instructor_seleccionado, materia_num_seleccionada)

st.sidebar.markdown("---")
st.sidebar.info(f"📚 **Materia Vinculada:**\n{materia_detectada}")

alumnos_ficha = df_aprendices[df_aprendices["Ficha"] == ficha_seleccionada].reset_index(drop=True)
st.sidebar.markdown(f"**Total Aprendices Activos:** {len(alumnos_ficha)}")

# --- PESTAÑAS DE LA APLICACIÓN ---
tab1, tab2, tab3 = st.tabs(["📋 Llamado a Lista", "📝 Evaluar Competencia", "📈 Historial y Reportes"])

# PESTAÑA 1: LLAMADO A LISTA (OPTIMIZADA CON st.form)
with tab1:
    st.header(f"📋 Control de Asistencia")
    st.subheader(f"Ficha: {ficha_seleccionada} | Instructor: {instructor_seleccionado}")
    st.caption(f"📖 Asignatura activa: {materia_detectada}")
    
    fecha_asistencia = st.date_input("Fecha del llamado a lista:", datetime.now())
    
    if alumnos_ficha.empty:
        st.warning(f"No hay aprendices activos asignados a la ficha {ficha_seleccionada}.")
    else:
        # Se crea el formulario dinámico para congelar las recargas intermitentes
        with st.form(key=f"formulario_asistencia_{ficha_seleccionada}"):
            st.markdown("Marque la casilla si el aprendiz se encuentra **Presente**:")
            
            asistencia_dict = {}
            col_doc, col_nom, col_estado = st.columns([2, 5, 2])
            col_doc.markdown("**Documento**")
            col_nom.markdown("**Nombre Completo**")
            col_estado.markdown("**¿Presente?**")
            st.markdown("---")
            
            # Renderizado de los alumnos
            for idx, row in alumnos_ficha.iterrows():
                c1, c2, c3 = st.columns([2, 5, 2])
                c1.text(row["Documento"])
                c2.text(row["Nombre Completo"])
                # Aquí el usuario interactúa libremente y rápido sin recargas
                asistencia_dict[idx] = c3.checkbox("Presente", value=True, key=f"check_{ficha_seleccionada}_{idx}")
            
            st.markdown("---")
            # El botón de guardar del formulario envía todo el paquete completo al final
            boton_guardar = st.form_submit_button("💾 Guardar Lista Completa", type="primary")
            
        if boton_guardar:
            registros = []
            for idx, row in alumnos_ficha.iterrows():
                estado = "Presente" if asistencia_dict[idx] else "Falta"
                registros.append({
                    "Fecha": fecha_asistencia,
                    "Ficha": ficha_seleccionada,
                    "Instructor": instructor_seleccionado,
                    "Materia_Num": materia_num_seleccionada,
                    "Materia_Nombre": materia_detectada,
                    "Documento": row["Documento"],
                    "Nombre": row["Nombre Completo"],
                    "Asistencia": estado
                })
            
            df_nuevo = pd.DataFrame(registros)
            file_exists = os.path.exists("asistencia_guardada.csv")
            df_nuevo.to_csv("asistencia_guardada.csv", mode='a', header=not file_exists, index=False)
            st.success(f"¡Asistencia de '{materia_detectada}' guardada exitosamente!")

# PESTAÑA 2: EVALUACIÓN (TAMBIÉN OPTIMIZADA CON st.form)
with tab2:
    st.header(f"📝 Registro de Juicios Evaluativos")
    st.markdown(f"**Ficha:** {ficha_seleccionada} | **Evaluación para:** {materia_detectada}")
    fecha_evaluacion = st.date_input("Fecha de Evaluación:", datetime.now())
    
    if alumnos_ficha.empty:
        st.warning("No hay alumnos para evaluar.")
    else:
        with st.form(key=f"formulario_evaluacion_{ficha_seleccionada}"):
            eval_dict = {}
            obs_dict = {}
            col_e_nom, col_e_cal, col_e_obs = st.columns([4, 2, 4])
            col_e_nom.markdown("**Aprendiz**")
            col_e_cal.markdown("**Juicio (A/D)**")
            col_e_obs.markdown("**Observaciones**")
            st.markdown("---")
            
            for idx, row in alumnos_ficha.iterrows():
                c1, c2, c3 = st.columns([4, 2, 4])
                c1.text(row["Nombre Completo"])
                eval_dict[idx] = c2.selectbox("Nota", ["A", "D"], key=f"eval_{ficha_seleccionada}_{idx}", label_visibility="collapsed")
                obs_dict[idx] = c3.text_input("Obs", placeholder="Alcanza el RA", key=f"obs_{ficha_seleccionada}_{idx}", label_visibility="collapsed")
                
            st.markdown("---")
            boton_guardar_eval = st.form_submit_button("💾 Guardar Evaluaciones", type="primary")
            
        if boton_guardar_eval:
            registros_eval = []
            for idx, row in alumnos_ficha.iterrows():
                registros_eval.append({
                    "Fecha": fecha_evaluacion,
                    "Ficha": ficha_seleccionada,
                    "Instructor": instructor_seleccionado,
                    "Materia": materia_detectada,
                    "Documento": row["Documento"],
                    "Nombre": row["Nombre Completo"],
                    "Evaluación (A/D)": eval_dict[idx],
                    "Observaciones": obs_dict[idx]
                })
            pd.DataFrame(registros_eval).to_csv("evaluaciones_guardadas.csv", mode='a', header=False, index=False)
            st.success("Evaluaciones guardadas con éxito.")

# PESTAÑA 3: REPORTES CONSOLIDADOS
with tab3:
    st.header("📈 Historial de Registros")
    sub_tab1, sub_tab2 = st.tabs(["Histórico de Asistencias", "Histórico de Notas"])
    
    with sub_tab1:
        if os.path.exists("asistencia_guardada.csv"):
            df_asist_hist = pd.read_csv("asistencia_guardada.csv")
            df_filtrado_asist = df_asist_hist[df_asist_hist["Ficha"].astype(str) == str(ficha_seleccionada)]
            if not df_filtrado_asist.empty:
                st.dataframe(df_filtrado_asist, use_container_width=True)
            else:
                st.info("No hay registros de asistencia para esta ficha.")"  Requiero ahora abrir otra ventana que me llene la información de la Hoja "Cabezote" Puede ser que se haga un archivo nuevo y que también se haga un archivo de la hoja "Listado de aprendices"
