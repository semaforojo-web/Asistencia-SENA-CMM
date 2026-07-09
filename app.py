import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuración de la página para entorno móvil y de escritorio
st.set_page_config(page_title="Control de Asistencia y Evaluación - SENA", layout="wide")

# Nombre del archivo original de Excel en tu repositorio de GitHub
DB_FILE = "Reporte de Asistencia.xlsx"

def obtener_instructores_y_contraseñas():
    """
    Lee los instructores (Columna A) y sus contraseñas (Columna B) 
    desde la hoja 'Listado de instructores'
    """
    dict_usuarios = {}
    lista_instructores = []
    
    if os.path.exists(DB_FILE):
        try:
            df_inst = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None)
            
            # Recorremos el archivo fila por fila
            for idx, row in df_inst.iterrows():
                nombre = str(row[0]).strip() if pd.notna(row[0]) else ""
                # Si no hay contraseña asignada en la columna B, por defecto se usa "SENA2026"
                password = str(row[1]).strip() if df_inst.shape[1] > 1 and pd.notna(row[1]) else "SENA2026"
                
                # Filtrar posibles encabezados o celdas vacías
                if nombre.upper() != "INSTRUCTOR" and nombre != "" and nombre.upper() != "NAN":
                    dict_usuarios[nombre] = password
                    lista_instructores.append(nombre)
                    
            return sorted(lista_instructores), dict_usuarios
        except Exception:
            return ["Falta hoja 'Listado de instructores'"], {}
            
    return ["No Detectado"], {}

def cargar_datos():
    """Carga y procesa el listado de aprendices desde la hoja correspondiente"""
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None)
            if df.empty:
                return pd.DataFrame(columns=["Grupo", "Documento", "Nombre Completo"])
                
            df_procesado = pd.DataFrame()
            
            # Columna E (Ficha -> Ahora Grupo) -> Posición 4
            df_procesado["Grupo"] = df.iloc[:, 4].astype(str).str.strip()
            # Columna L (Documento) -> Posición 11
            df_procesado["Documento"] = df.iloc[:, 11].fillna("S/D").astype(str).str.strip()
            # Columnas M a O (Nombres y Apellidos) -> Posiciones 12 a 14
            nombres_completos = df.iloc[:, 12:15].fillna("").astype(str)
            df_procesado["Nombre Completo"] = nombres_completos.agg(' '.join, axis=1).str.replace(r'\s+', ' ', regex=True).str.strip()
            
            # FILTRADO POR COLUMNA P (Posición 15) - Estados de deserción o retiro
            df_procesado["Estado"] = df.iloc[:, 15].fillna("").astype(str).str.upper().str.strip()
            terminos_excluir = "CANCELADO|RETIRO VOLUNTARIO|TRASLADO"
            df_procesado = df_procesado[~df_procesado["Estado"].str.contains(terminos_excluir, regex=True)]
            
            # Asegurar que el grupo sea numérico válido
            df_procesado = df_procesado[df_procesado["Grupo"].str.contains(r'^\d+$', na=False)]
            df_procesado["Nombre Completo"] = df_procesado["Nombre Completo"].replace("", "Aprendiz sin nombre registrado")
            
            return df_procesado[["Grupo", "Documento", "Nombre Completo"]].reset_index(drop=True)
        except Exception as e:
            st.error(f"Error al filtrar listado de aprendices: {e}")
            
    data = {"Grupo": ["Error"], "Documento": ["0"], "Nombre Completo": ["ARCHIVO EXCEL NO DETECTADO"]}
    return pd.DataFrame(data)

def obtener_trimestres_disponibles(grupo, instructor):
    """Busca los trimestres vinculados a un grupo e instructor en el Cabezote"""
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            
            filtro = (df_cab[6] == str(grupo)) & (df_cab[5] == str(instructor).upper())
            resultado = df_cab[filtro]
            
            if not resultado.empty and resultado.shape[1] > 47:
                trimestres = resultado.iloc[:, 47].dropna().astype(str).str.strip().unique().tolist()
                trimestres_validos = sorted([t for t in trimestres if t != "" and t.upper() != "NAN" and t.upper() != "TRIMESTRE"])
                if trimestres_validos:
                    return trimestres_validos
        except Exception:
            pass
    return ["Sin trimestres detectados"]

def obtener_materias_disponibles(grupo, instructor, trimestre):
    """Obtiene los números de asignación cargados para la combinación seleccionada"""
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            df_cab[47] = df_cab[47].astype(str).str.strip() if df_cab.shape[1] > 47 else ""
            
            filtro = (df_cab[6] == str(grupo)) & \
                     (df_cab[5] == str(instructor).upper()) & \
                     (df_cab[47] == str(trimestre))
            resultado = df_cab[filtro]
            
            if not resultado.empty:
                materias = resultado.iloc[:, 3].dropna().astype(str).str.strip().unique().tolist()
                materias_validas = sorted([m for m in materias if m != "" and m.upper() != "MATERIA" and m.upper() != "NAN"])
                if materias_validas:
                    return materias_validas
        except Exception:
            pass
    return ["1", "2", "3"]

def filtrar_materia_final(grupo, instructor, trimestre, asignacion_num):
    """Retorna el nombre largo de la asignatura basándose en el número de asignación seleccionado"""
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[3] = df_cab[3].astype(str).str.strip()
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            df_cab[47] = df_cab[47].astype(str).str.strip() if df_cab.shape[1] > 47 else ""
            
            filtro = (df_cab[6] == str(grupo)) & \
                     (df_cab[5] == str(instructor).upper()) & \
                     (df_cab[47] == str(trimestre)) & \
                     (df_cab[3] == str(asignacion_num))
                     
            resultado = df_cab[filtro]
            
            if not resultado.empty:
                materia_texto = resultado.iloc[0, 10] if resultado.shape[1] > 10 else resultado.iloc[0, 3]
                if pd.notna(materia_texto) and str(materia_texto).strip() != "":
                    return str(materia_texto).strip()
            return f"Asignación {asignacion_num} (Sin descripción en Cabezote)"
        except Exception as e:
            return f"Error: {e}"
    return "Archivo Excel no encontrado"

# --- MANEJO DEL ESTADO DE SESIÓN (LOGIN) ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "instructor_logueado" not in st.session_state:
    st.session_state["instructor_logueado"] = None

# Cargar la lista de usuarios y contraseñas desde el Excel
lista_instructores_login, mapeo_credenciales = obtener_instructores_y_contraseñas()

# --- PANTALLA DE INGRESO (LOGIN) ---
if not st.session_state["autenticado"]:
    st.markdown("<h2 style='text-align: center;'>🔐 Acceso al Sistema de Asistencia - SENA</h2>", unsafe_allow_html=True)
    
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.form("form_login"):
                instructor_input = st.selectbox("Seleccione su Nombre de Instructor:", lista_instructores_login)
                password_input = st.text_input("Introduzca su contraseña personal:", type="password")
                
                boton_ingresar = st.form_submit_button("🚀 Ingresar a la Aplicación", use_container_width=True)
                
            if boton_ingresar:
                if instructor_input not in ["No Detectado", "Falta hoja 'Listado de instructores'"]:
                    contraseña_esperada = mapeo_credenciales.get(instructor_input, "SENA2026")
                    
                    if password_input == contraseña_esperada:
                        st.session_state["autenticado"] = True
                        st.session_state["instructor_logueado"] = instructor_input
                        st.success("¡Acceso concedido!")
                        st.rerun()
                    else:
                        st.error("Contraseña incorrecta para el instructor seleccionado. Inténtelo de nuevo.")
                else:
                    st.error("Por favor, verifique la estructura de instructores en su archivo Excel.")
    st.stop()

# --- EJECUCIÓN PRINCIPAL ---
df_aprendices = cargar_datos()
instructor_seleccionado = st.session_state["instructor_logueado"]

if not os.path.exists("asistencia_guardada.csv"):
    pd.DataFrame(columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Asignacion_Num", "Materia_Nombre", "Documento", "Nombre", "Asistencia"]).to_csv("asistencia_guardada.csv", index=False)

if not os.path.exists("evaluaciones_guardadas.csv"):
    pd.DataFrame(columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Materia", "Documento", "Nombre", "Evaluación (A/D)", "Observaciones"]).to_csv("evaluaciones_guardadas.csv", index=False)

# --- INTERFAZ SIDEBAR ---
st.title("📊 Dashboard de Gestión de Ambientes de Formación - SENA")

if st.sidebar.button("🔒 Cerrar Sesión"):
    st.session_state["autenticado"] = False
    st.session_state["instructor_logueado"] = None
    st.rerun()

st.sidebar.markdown(f"👤 **Instructor activo:**\n`{instructor_seleccionado}`")
st.sidebar.markdown("---")
st.sidebar.header("⚙️ Filtros de Planificación")

lista_grupos = sorted(df_aprendices["Grupo"].dropna().unique())
grupo_seleccionado = st.sidebar.selectbox("1. Seleccione el Grupo (Columna G):", lista_grupos if lista_grupos else ["Error"])

st.sidebar.markdown(f"**2. Seleccione el Instructor (Columna F):**\n`{instructor_seleccionado}`")

lista_trimestres_dinamicos = obtener_trimestres_disponibles(grupo_seleccionado, instructor_seleccionado)
trimestre_seleccionado = st.sidebar.selectbox("3. Seleccione el Trimestre (Columna AV):", lista_trimestres_dinamicos)

lista_asignaciones_dinamicas = obtener_materias_disponibles(grupo_seleccionado, instructor_seleccionado, trimestre_seleccionado)
asignacion_num_seleccionada = st.sidebar.selectbox("4. Seleccione Asignación (Columna D):", lista_asignaciones_dinamicas)

materia_detectada = filtrar_materia_final(grupo_seleccionado, instructor_seleccionado, trimestre_seleccionado, asignacion_num_seleccionada)

st.sidebar.markdown("---")
st.sidebar.info(f"📚 **Materia Vinculada:**\n{materia_detectada}")

alumnos_grupo = df_aprendices[df_aprendices["Grupo"] == grupo_seleccionado].reset_index(drop=True)
st.sidebar.markdown(f"**Total Aprendices Activos:** {len(alumnos_grupo)}")

# --- PESTAÑAS ---
tab1, tab2, tab3, tab4 = st.tabs(["📋 Llamado a Lista", "📝 Evaluar Competencia", "📈 Historial y Reportes", "📂 Alimentar y Cargar Bases"])

# PESTAÑA 1, 2 y 3 se mantienen iguales (Compactadas para mantener foco en el cambio)
with tab1:
    st.header(f"📋 Control de Asistencia")
    st.subheader(f"Grupo: {grupo_seleccionado} | Trimestre: {trimestre_seleccionado}")
    st.caption(f"📖 Asignatura: {materia_detectada} (Asignación {asignacion_num_seleccionada})")
    fecha_asistencia = st.date_input("Fecha del llamado a lista:", datetime.now())
    if alumnos_grupo.empty or grupo_seleccionado == "Error":
        st.warning(f"No hay aprendices activos asignados al grupo.")
    else:
        with st.form(key=f"formulario_asistencia_{grupo_seleccionado}"):
            asistencia_dict = {}
            col_doc, col_nom, col_estado = st.columns([2, 5, 2])
            col_doc.markdown("**Documento**")
            col_nom.markdown("**Nombre Completo**")
            col_estado.markdown("**¿Presente?**")
            for idx, row in alumnos_grupo.iterrows():
                c1, c2, c3 = st.columns([2, 5, 2])
                c1.text(row["Documento"])
                c2.text(row["Nombre Completo"])
                asistencia_dict[idx] = c3.checkbox("Presente", value=True, key=f"check_{grupo_seleccionado}_{idx}")
            boton_guardar = st.form_submit_button("💾 Guardar Lista Completa", type="primary")
        if boton_guardar:
            registros = []
            for idx, row in alumnos_grupo.iterrows():
                estado = "Presente" if asistencia_dict[idx] else "Falta"
                registros.append({"Fecha": fecha_asistencia, "Grupo": grupo_seleccionado, "Instructor": instructor_seleccionado, "Trimestre": trimestre_seleccionado, "Asignacion_Num": asignacion_num_seleccionada, "Materia_Nombre": materia_detectada, "Documento": row["Documento"], "Nombre": row["Nombre Completo"], "Asistencia": estado})
            pd.DataFrame(registros).to_csv("asistencia_guardada.csv", mode='a', header=False, index=False)
            st.success(f"¡Asistencia guardada!")
            st.rerun()

with tab2:
    st.header(f"📝 Registro de Juicios Evaluativos")
    fecha_evaluacion = st.date_input("Fecha de Evaluación:", datetime.now())
    if not alumnos_grupo.empty and grupo_seleccionado != "Error":
        with st.form(key=f"formulario_evaluacion_{grupo_seleccionado}"):
            eval_dict, obs_dict = {}, {}
            for idx, row in alumnos_grupo.iterrows():
                c1, c2, c3 = st.columns([4, 2, 4])
                c1.text(row["Nombre Completo"])
                eval_dict[idx] = c2.selectbox("Nota", ["A", "D"], key=f"eval_{grupo_seleccionado}_{idx}", label_visibility="collapsed")
                obs_dict[idx] = c3.text_input("Obs", placeholder="Alcanza el RA", key=f"obs_{grupo_seleccionado}_{idx}", label_visibility="collapsed")
            boton_guardar_eval = st.form_submit_button("💾 Guardar Evaluaciones", type="primary")
        if boton_guardar_eval:
            registros_eval = []
            for idx, row in alumnos_grupo.iterrows():
                registros_eval.append({"Fecha": fecha_evaluacion, "Grupo": grupo_seleccionado, "Instructor": instructor_seleccionado, "Trimestre": trimestre_seleccionado, "Materia": materia_detectada, "Documento": row["Documento"], "Nombre": row["Nombre Completo"], "Evaluación (A/D)": eval_dict[idx], "Observaciones": obs_dict[idx]})
            pd.DataFrame(registros_eval).to_csv("evaluaciones_guardadas.csv", mode='a', header=False, index=False)
            st.success("Evaluaciones guardadas.")

with tab3:
    st.header("📈 Historial de Registros")
    sub_tab1, sub_tab2 = st.tabs(["Histórico de Asistencias", "Histórico de Notas"])
    with sub_tab1:
        if os.path.exists("asistencia_guardada.csv"):
            df_asist_hist = pd.read_csv("asistencia_guardada.csv")
            if "Ficha" in df_asist_hist.columns: df_asist_hist = df_asist_hist.rename(columns={"Ficha": "Grupo", "Materia_Num": "Asignacion_Num"})
            if "Grupo" in df_asist_hist.columns:
                df_filtrado_asist = df_asist_hist[(df_asist_hist["Grupo"].astype(str) == str(grupo_seleccionado)) & (df_asist_hist["Instructor"].astype(str).str.upper() == str(instructor_seleccionado).upper())]
                st.dataframe(df_filtrado_asist, use_container_width=True)
    with sub_tab2:
        if os.path.exists("evaluaciones_guardadas.csv"):
            df_eval_hist = pd.read_csv("evaluaciones_guardadas.csv")
            if "Ficha" in df_eval_hist.columns: df_eval_hist = df_eval_hist.rename(columns={"Ficha": "Grupo"})
            if "Grupo" in df_eval_hist.columns:
                df_filtrado_eval = df_eval_hist[(df_eval_hist["Grupo"].astype(str) == str(grupo_seleccionado)) & (df_eval_hist["Instructor"].astype(str).str.upper() == str(instructor_seleccionado).upper())]
                st.dataframe(df_filtrado_eval, use_container_width=True)

# PESTAÑA 4: GESTIÓN DE BASES DE DATOS (NUEVA ESTRUCTURA AMPLIADA G HASTA AV)
with tab4:
    st.header("📂 GESTIÓN Y ALIMENTACIÓN DESDE LA COLUMNA G HASTA LA AV")
    
    opcion_carga = st.radio("Seleccione el método para gestionar datos:", ["✍️ Alimentar Cabezote Directamente (Formulario)", "📁 Subir Archivos Completos (.xlsx)"])
    
    if opcion_carga == "✍️ Alimentar Cabezote Directamente (Formulario)":
        st.subheader("📝 Formulario Maestro del Ambiente de Formación")
        
        with st.form("form_registro_directo_cabezote_completo"):
            # Fila 1: Datos Base (A, F, G, D)
            c1, c2, c3, c4 = st.columns(4)
            input_grupo = c1.text_input("Número de Grupo (Columna G / Posición 6):", placeholder="Ej: 3141501")
            input_instructor = c2.text_input("Nombre del Instructor (Columna F / Posición 5):", value=instructor_seleccionado)
            input_asignacion_num = c3.selectbox("Número de asignación (Columna D / Posición 3):", ["1", "2", "3"])
            input_materia_nombre = c4.text_input("Resultados de Aprendizaje (Columna K / Posición 10):", placeholder="Ej: Mantenimiento")
            
            # Fila 2: Información del Proyecto
            st.markdown("##### 🚀 Planificación Estratégica del Proyecto")
            c5, c6, c7 = st.columns(3)
            input_fase = c5.text_input("Fase del proyecto (Columna H / Posición 7):")
            input_actividades = c6.text_area("Actividades del Proyecto (Columna I / Posición 8):", height=68)
            input_competencia = c7.text_area("Competencia (Columna J / Posición 9):", height=68)
            
            # Fila 3: Resultados, Horas y Tiempos
            c8, c9, c10, c11 = st.columns([3, 1, 1, 1])
            input_ra = c8.text_input("Resultados de Aprendizaje (Columna L / Posición 11):")
            input_linea = c9.text_input("Línea (Columna M / Posición 12):")
            input_horas = c10.text_input("Número de Horas (Columna N / Posición 13):")
            input_fecha_ini = c11.text_input("Fecha de inicio (Columna O / Posición 14):", placeholder="DD/MM/AAAA")
            
            # Fila 4: Ambientes y Horarios
            st.markdown("##### 🏫 Datos Locales del Ambiente y Jornada")
            c12, c13, c14, c15 = st.columns(4)
            input_ambiente = c12.text_input("Ambiente (Columna Q / Posición 16):")
            input_dia = c13.text_input("Día (Columna R / Posición 17):")
            input_horario = c14.text_input("Horario (Columna S / Posición 18):")
            input_jornada = c15.text_input("Jornada (Columna T / Posición 19):")
            
            # Fila 5: Evidencias
            st.markdown("##### 📑 Evidencias del Proceso (1 al 5)")
            col_ev1, col_ev2, col_ev3, col_ev4, col_ev5 = st.columns(5)
            input_ev1 = col_ev1.text_input("Evidencia 1 (Pos 20):")
            input_ev2 = col_ev2.text_input("Evidencia 2 (Pos 21):")
            input_ev3 = col_ev3.text_input("Evidencia 3 (Pos 22):")
            input_ev4 = col_ev4.text_input("Evidencia 4 (Pos 23):")
            input_ev5 = col_ev5.text_input("Evidencia 5 (Pos 24):")
            
            # Desplegables para descripciones prácticas (1 a 11)
            with st.expander("🛠️ Descripciones Prácticas (1 al 11) - Posiciones 25 a 35"):
                cp1, cp2, cp3, cp4 = st.columns(4)
                input_p1 = cp1.text_input("Descripción práctica 1:")
                input_p2 = cp2.text_input("Descripción práctica 2:")
                input_p3 = cp3.text_input("Descripción práctica 3:")
                input_p4 = cp4.text_input("Descripción práctica 4:")
                
                cp5, cp6, cp7, cp8 = st.columns(4)
                input_p5 = cp5.text_input("Descripción práctica 5:")
                input_p6 = cp6.text_input("Descripción práctica 6:")
                input_p7 = cp7.text_input("Descripción práctica 7:")
                input_p8 = cp8.text_input("Descripción práctica 8:")
                
                cp9, cp10, cp11 = st.columns(3)
                input_p9 = cp9.text_input("Descripción práctica 9:")
                input_p10 = cp10.text_input("Descripción práctica 10:")
                input_p11 = cp11.text_input("Descripción práctica 11:")
                
            # Desplegables para nombres de sesiones (1 a 11) y Trimestre
            with st.expander("📅 Nombres de la Sesión (1 al 11) y Observaciones Finales"):
                cs1, cs2, cs3, cs4 = st.columns(4)
                input_s1 = cs1.text_input("Nombre de la sesión 1:")
                input_s2 = cs2.text_input("Nombre de la sesión 2:")
                input_s3 = cs3.text_input("Nombre de la sesión 3:")
                input_s4 = cs4.text_input("Nombre de la sesión 4:")
                
                cs5, cs6, cs7, cs8 = st.columns(4)
                input_s5 = cs5.text_input("Nombre de la sesión 5:")
                input_s6 = cs6.text_input("Nombre de la sesión 6:")
                input_s7 = cs7.text_input("Nombre de la sesión 7:")
                input_s8 = cs8.text_input("Nombre de la sesión 8:")
                
                cs9, cs10, cs11 = st.columns(3)
                input_s9 = cs9.text_input("Nombre de la sesión 9:")
                input_s10 = cs10.text_input("Nombre de la sesión 10:")
                input_s11 = cs11.text_input("Nombre de la sesión 11:")
                
                st.markdown("---")
                c_f1, c_f2 = st.columns(2)
                input_trimestre = c_f1.text_input("Trimestre reportado (Columna AV / Posición 47):", placeholder="Ej: Trimestre 1")
                input_observaciones = c_f2.text_input("Observaciones (Columna AW / Posición 48):")
            
            boton_agregar_cab = st.form_submit_button("💾 Insertar Registro Estructurado Completo", type="primary")
            
        if boton_agregar_cab:
            if input_grupo and input_instructor and input_trimestre:
                try:
                    df_cab_existente = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    df_apr_existente = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    
                    # Definimos el ancho exacto del Cabezote (49 celdas indexadas de 0 a 48)
                    ancho_columnas = 49
                    nueva_fila = [""] * ancho_columnas
                    # Mapeo Exacto de los Índices según tus especificaciones
                    nueva_fila[3] = str(input_asignacion_num).strip()          # Columna D
                    nueva_fila[5] = str(input_instructor).strip().upper()       # Columna F
                    nueva_fila[6] = str(input_grupo).strip()                    # Columna G (GRUPO)
                    nueva_fila[7] = str(input_fase).strip()                     # Columna H (Fase)
                    nueva_fila[8] = str(input_actividades).strip()              # Columna I (Actividades)
                    nueva_fila[9] = str(input_competencia).strip()              # Columna J (Competencia)
                    nueva_fila[10] = str(input_materia_nombre).strip()          # Columna K
                    nueva_fila[11] = str(input_ra).strip()                      # Columna L (RA)
                    nueva_fila[12] = str(input_linea).strip()                   # Columna M (Línea)
                    nueva_fila[13] = str(input_horas).strip()                   # Columna N (Horas)
                    nueva_fila[14] = str(input_fecha_ini).strip()               # Columna O (Fecha inicio)
                    # Columna P queda libre o vacía ("")
                    nueva_fila[16] = str(input_ambiente).strip()                # Columna Q (Ambiente)
                    nueva_fila[17] = str(input_dia).strip()                     # Columna R (Día)
                    nueva_fila[18] = str(input_horario).strip()                 # Columna S (Horario)
                    nueva_fila[19] = str(input_jornada).strip()                 # Columna T (Jornada)
                    
                    # Evidencias 1 a 5 (Posiciones 20 a 24)
                    nueva_fila[20] = str(input_ev1).strip()
                    nueva_fila[21] = str(input_ev2).strip()
                    nueva_fila[22] = str(input_ev3).strip()
                    nueva_fila[23] = str(input_ev4).strip()
                    nueva_fila[24] = str(input_ev5).strip()
                    
                    # Descripciones Prácticas 1 a 11 (Posiciones 25 a 35)
                    nueva_fila[25] = str(input_p1).strip()
                    nueva_fila[26] = str(input_p2).strip()
                    nueva_fila[27] = str(input_p3).strip()
                    nueva_fila[28] = str(input_p4).strip()
                    nueva_fila[29] = str(input_p5).strip()
                    nueva_fila[30] = str(input_p6).strip()
                    nueva_fila[31] = str(input_p7).strip()
                    nueva_fila[32] = str(input_p8).strip()
                    nueva_fila[33] = str(input_p9).strip()
                    nueva_fila[34] = str(input_p10).strip()
                    nueva_fila[35] = str(input_p11).strip()
                    
                    # Nombres de las Sesiones 1 a 11 (Posiciones 36 a 46)
                    nueva_fila[36] = str(input_s1).strip()
                    nueva_fila[37] = str(input_s2).strip()
                    nueva_fila[38] = str(input_s3).strip()
                    nueva_fila[39] = str(input_s4).strip()
                    nueva_fila[40] = str(input_s5).strip()
                    nueva_fila[41] = str(input_s6).strip()
                    nueva_fila[42] = str(input_s7).strip()
                    nueva_fila[43] = str(input_s8).strip()
                    nueva_fila[44] = str(input_s9).strip()
                    nueva_fila[45] = str(input_s10).strip()
                    nueva_fila[46] = str(input_s11).strip()
                    
                    # Cierre
                    nueva_fila[47] = str(input_trimestre).strip()               # Columna AV (Trimestre)
                    nueva_fila[48] = str(input_observaciones).strip()           # Columna AW (Observaciones)
                    
                    df_nueva_fila = pd.DataFrame([nueva_fila])
                    
                    if not df_cab_existente.empty:
                        # Forzamos que coincidan las estructuras rellenando columnas faltantes
                        if df_cab_existente.shape[1] < ancho_columnas:
                            for col_idx in range(df_cab_existente.shape[1], ancho_columnas):
                                df_cab_existente[col_idx] = ""
                        df_nueva_fila.columns = df_cab_existente.columns[:ancho_columnas]
                        df_cab_final = pd.concat([df_cab_existente, df_nueva_fila], ignore_index=True)
                    else:
                        df_cab_final = df_nueva_fila
                    
                    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                        df_cab_final.to_excel(writer, sheet_name="Cabezote", index=False, header=False)
                        if not df_apr_existente.empty:
                            df_apr_existente.to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                            
                    st.success("¡Ambiente maestro e información de sesiones inyectados correctamente!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al escribir en el Excel: {e}")
            else:
                st.warning("El Grupo, Instructor y Trimestre son obligatorios para indexar el registro.")

    elif opcion_carga == "📁 Subir Archivos Completos (.xlsx)":
        # Se mantiene la carga manual intacta para casos de sobreescritura total
        st.markdown("Sube las hojas de cálculo por separado para realizar un empaquetado general.")
        file_cabezote = st.file_uploader("Subir archivo para Cabezote (.xlsx)", type=["xlsx"])
        file_aprendices = st.file_uploader("Subir archivo para Aprendices (.xlsx)", type=["xlsx"])
        if st.button("🧩 Procesar e Integrar Base de Datos Maestro"):
            if file_cabezote and file_aprendices:
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    pd.read_excel(file_cabezote, header=None).to_excel(writer, sheet_name="Cabezote", index=False, header=False)
                    pd.read_excel(file_aprendices, header=None).to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                st.success("Configurado con éxito.")
                st.rerun()
