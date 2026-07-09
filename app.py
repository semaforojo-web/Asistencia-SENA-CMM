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

# --- EJECUCIÓN PRINCIPAL (SÓLO SI PASÓ EL LOGIN) ---
df_aprendices = cargar_datos()
instructor_seleccionado = st.session_state["instructor_logueado"]

# Creación de históricos locales si no existen (Ya estructurados con Grupo y Asignacion_Num)
if not os.path.exists("asistencia_guardada.csv"):
    pd.DataFrame(columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Asignacion_Num", "Materia_Nombre", "Documento", "Nombre", "Asistencia"]).to_csv("asistencia_guardada.csv", index=False)

if not os.path.exists("evaluaciones_guardadas.csv"):
    pd.DataFrame(columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Materia", "Documento", "Nombre", "Evaluación (A/D)", "Observaciones"]).to_csv("evaluaciones_guardadas.csv", index=False)

# --- INTERFAZ SIDEBAR (BARRA LATERAL) ---
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

# --- PESTAÑAS DE LA APLICACIÓN ---
tab1, tab2, tab3, tab4 = st.tabs(["📋 Llamado a Lista", "📝 Evaluar Competencia", "📈 Historial y Reportes", "📂 Alimentar y Cargar Bases"])

# PESTAÑA 1: LLAMADO A LISTA
with tab1:
    st.header(f"📋 Control de Asistencia")
    st.subheader(f"Grupo: {grupo_seleccionado} | Trimestre: {trimestre_seleccionado}")
    st.caption(f"📖 Asignatura: {materia_detectada} (Asignación {asignacion_num_seleccionada})")
    
    fecha_asistencia = st.date_input("Fecha del llamado a lista:", datetime.now())
    
    if alumnos_grupo.empty or grupo_seleccionado == "Error":
        st.warning(f"No hay aprendices activos asignados al grupo seleccionado o el archivo maestro está vacío.")
    else:
        with st.form(key=f"formulario_asistencia_{grupo_seleccionado}"):
            st.markdown("Marque la casilla si el aprendiz se encuentra **Presente**:")
            
            asistencia_dict = {}
            col_doc, col_nom, col_estado = st.columns([2, 5, 2])
            col_doc.markdown("**Documento**")
            col_nom.markdown("**Nombre Completo**")
            col_estado.markdown("**¿Presente?**")
            st.markdown("---")
            
            for idx, row in alumnos_grupo.iterrows():
                c1, c2, c3 = st.columns([2, 5, 2])
                c1.text(row["Documento"])
                c2.text(row["Nombre Completo"])
                asistencia_dict[idx] = c3.checkbox("Presente", value=True, key=f"check_{grupo_seleccionado}_{idx}")
            
            st.markdown("---")
            boton_guardar = st.form_submit_button("💾 Guardar Lista Completa", type="primary")
            
        if boton_guardar:
            registros = []
            for idx, row in alumnos_grupo.iterrows():
                estado = "Presente" if asistencia_dict[idx] else "Falta"
                registros.append({
                    "Fecha": fecha_asistencia,
                    "Grupo": grupo_seleccionado,
                    "Instructor": instructor_seleccionado,
                    "Trimestre": trimestre_seleccionado,
                    "Asignacion_Num": asignacion_num_seleccionada,
                    "Materia_Nombre": materia_detectada,
                    "Documento": row["Documento"],
                    "Nombre": row["Nombre Completo"],
                    "Asistencia": estado
                })
            
            df_nuevo = pd.DataFrame(registros)
            df_nuevo.to_csv("asistencia_guardada.csv", mode='a', header=False, index=False)
            st.success(f"¡Asistencia de '{materia_detectada}' guardada exitosamente!")
            st.rerun()

# PESTAÑA 2: EVALUACIÓN
with tab2:
    st.header(f"📝 Registro de Juicios Evaluativos")
    st.markdown(f"**Grupo:** {grupo_seleccionado} | **Trimestre:** {trimestre_seleccionado} | **Asignación {asignacion_num_seleccionada}:** {materia_detectada}")
    fecha_evaluacion = st.date_input("Fecha de Evaluación:", datetime.now())
    
    if alumnos_grupo.empty or grupo_seleccionado == "Error":
        st.warning("No hay alumnos para evaluar.")
    else:
        with st.form(key=f"formulario_evaluacion_{grupo_seleccionado}"):
            eval_dict = {}
            obs_dict = {}
            col_e_nom, col_e_cal, col_e_obs = st.columns([4, 2, 4])
            col_e_nom.markdown("**Aprendiz**")
            col_e_cal.markdown("**Juicio (A/D)**")
            col_e_obs.markdown("**Observaciones**")
            st.markdown("---")
            
            for idx, row in alumnos_grupo.iterrows():
                c1, c2, c3 = st.columns([4, 2, 4])
                c1.text(row["Nombre Completo"])
                eval_dict[idx] = c2.selectbox("Nota", ["A", "D"], key=f"eval_{grupo_seleccionado}_{idx}", label_visibility="collapsed")
                obs_dict[idx] = c3.text_input("Obs", placeholder="Alcanza el RA", key=f"obs_{grupo_seleccionado}_{idx}", label_visibility="collapsed")
                
            st.markdown("---")
            boton_guardar_eval = st.form_submit_button("💾 Guardar Evaluaciones", type="primary")
            
        if boton_guardar_eval:
            registros_eval = []
            for idx, row in alumnos_grupo.iterrows():
                registros_eval.append({
                    "Fecha": fecha_evaluacion,
                    "Grupo": grupo_seleccionado,
                    "Instructor": instructor_seleccionado,
                    "Trimestre": trimestre_seleccionado,
                    "Materia": materia_detectada,
                    "Documento": row["Documento"],
                    "Nombre": row["Nombre Completo"],
                    "Evaluación (A/D)": eval_dict[idx],
                    "Observaciones": obs_dict[idx]
                })
            pd.DataFrame(registros_eval).to_csv("evaluaciones_guardadas.csv", mode='a', header=False, index=False)
            st.success("Evaluaciones guardadas con éxito.")

# PESTAÑA 3: REPORTES CONSOLIDADOS (BLINDADA CONTRA EL KEYERROR)
with tab3:
    st.header("📈 Historial de Registros")
    sub_tab1, sub_tab2 = st.tabs(["Histórico de Asistencias", "Histórico de Notas"])
    
    with sub_tab1:
        if os.path.exists("asistencia_guardada.csv"):
            df_asist_hist = pd.read_csv("asistencia_guardada.csv")
            
            # Si el archivo viejo existe y contiene "Ficha", lo renombramos sobre la marcha para que no falle
            if "Ficha" in df_asist_hist.columns:
                df_asist_hist = df_asist_hist.rename(columns={"Ficha": "Grupo", "Materia_Num": "Asignacion_Num"})
            
            if "Grupo" in df_asist_hist.columns:
                df_filtrado_asist = df_asist_hist[(df_asist_hist["Grupo"].astype(str) == str(grupo_seleccionado)) & (df_asist_hist["Instructor"].astype(str).str.upper() == str(instructor_seleccionado).upper())]
                if not df_filtrado_asist.empty:
                    st.dataframe(df_filtrado_asist, use_container_width=True)
                else:
                    st.info("No hay registros de asistencia guardados por usted para este grupo.")
            else:
                st.info("Estructura de archivo no reconocida.")
                
    with sub_tab2:
        if os.path.exists("evaluaciones_guardadas.csv"):
            df_eval_hist = pd.read_csv("evaluaciones_guardadas.csv")
            
            # Lo mismo para evaluaciones
            if "Ficha" in df_eval_hist.columns:
                df_eval_hist = df_eval_hist.rename(columns={"Ficha": "Grupo"})
                
            if "Grupo" in df_eval_hist.columns:
                df_filtrado_eval = df_eval_hist[(df_eval_hist["Grupo"].astype(str) == str(grupo_seleccionado)) & (df_eval_hist["Instructor"].astype(str).str.upper() == str(instructor_seleccionado).upper())]
                if not df_filtrado_eval.empty:
                    st.dataframe(df_filtrado_eval, use_container_width=True)
                else:
                    st.info("No hay registros de evaluaciones guardados por usted para este grupo.")
            else:
                st.info("Estructura de archivo no reconocida.")

# PESTAÑA 4: GESTIÓN DE BASES DE DATOS
with tab4:
    st.header("📂 Gestión y Alimentación de la Base de Datos")
    
    opcion_carga = st.radio("Seleccione el método para gestionar datos:", ["✍️ Alimentar Cabezote Directamente (Formulario)", "📁 Subir Archivos Completos (.xlsx)"])
    
    if opcion_carga == "✍️ Alimentar Cabezote Directamente (Formulario)":
        st.subheader("📝 Agregar un nuevo registro a la hoja 'Cabezote'")
        st.markdown("Ingresa los datos del ambiente. Los campos restantes se generarán en blanco automáticamente.")
        
        with st.form("form_registro_directo_cabezote"):
            c1, c2 = st.columns(2)
            input_grupo = c1.text_input("Número de Grupo (Columna G / Posición 6):", placeholder="Ej: 2613452")
            input_instructor = c2.text_input("Nombre del Instructor (Columna F / Posición 5):", value=instructor_seleccionado)
            
            c3, c4 = st.columns(2)
            input_asignacion_num = c3.selectbox("Número de asignación (Columna D / Posición 3):", ["1", "2", "3"])
            input_trimestre = c4.text_input("Trimestre de la formación (Columna AV / Posición 47):", placeholder="Ej: Trimestre 3")
            
            input_materia_nombre = st.text_input("Nombre Detallado de la Asignatura (Columna K / Posición 10):", placeholder="Ej: Mantenimiento Correctivo de Sistemas Electromecánicos")
            
            boton_agregar_cab = st.form_submit_button("💾 Insertar Fila en Cabezote de Excel", type="primary")
            
        if boton_agregar_cab:
            if input_grupo and input_instructor and input_materia_nombre and input_trimestre:
                try:
                    df_cab_existente = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    df_apr_existente = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    
                    ancho_columnas = max(df_cab_existente.shape[1], 48) if not df_cab_existente.empty else 48
                    
                    nueva_fila = [""] * ancho_columnas
                    nueva_fila[3] = str(input_asignacion_num).strip()
                    nueva_fila[5] = str(input_instructor).strip().upper()
                    nueva_fila[6] = str(input_grupo).strip()
                    nueva_fila[10] = str(input_materia_nombre).strip()
                    nueva_fila[47] = str(input_trimestre).strip()
                    
                    df_nueva_fila = pd.DataFrame([nueva_fila])
                    
                    if not df_cab_existente.empty:
                        df_nueva_fila.columns = df_cab_existente.columns[:ancho_columnas] if df_cab_existente.shape[1] >= ancho_columnas else range(ancho_columnas)
                        df_cab_final = pd.concat([df_cab_existente, df_nueva_fila], ignore_index=True)
                    else:
                        df_cab_final = df_nueva_fila
                    
                    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                        df_cab_final.to_excel(writer, sheet_name="Cabezote", index=False, header=False)
                        
                        if not df_apr_existente.empty:
                            df_apr_existente.to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                        else:
                            base_aprendices = [""] * 16
                            base_aprendices[4] = "Grupo"
                            base_aprendices[11] = "Documento"
                            base_aprendices[12] = "Nombre"
                            base_aprendices[15] = "Estado"
                            pd.DataFrame([base_aprendices]).to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                            
                    st.success("¡Ambiente registrado e insertado con éxito en el archivo maestro de Excel!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al escribir de forma directa en el archivo Excel: {e}")
            else:
                st.warning("Por favor rellene todos los campos solicitados para estructurar la fila correctamente.")
                
        if os.path.exists(DB_FILE):
            try:
                st.markdown("### 📋 Vista Previa del Cabezote Guardado Real:")
                df_preview = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
                if df_preview.shape[1] > 47:
                    df_resumen = pd.DataFrame({
                        "Grupo": df_preview[6],
                        "Instructor": df_preview[5],
                        "Número Asignación": df_preview[3],
                        "Asignatura": df_preview[10],
                        "Trimestre": df_preview[47]
                    }).dropna(subset=["Grupo", "Instructor"], how="all")
                    st.dataframe(df_resumen, use_container_width=True)
            except Exception:
                pass

    elif opcion_carga == "📁 Subir Archivos Completos (.xlsx)":
        st.markdown("Sube las hojas de cálculo por separado para realizar un empaquetado general desde cero.")
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            st.subheader("1. Hoja Cabezote")
            file_cabezote = st.file_uploader("Subir archivo para Cabezote (.xlsx)", type=["xlsx"], key="up_cab_v2")
        with col_up2:
            st.subheader("2. Hoja Listado de Aprendices")
            file_aprendices = st.file_uploader("Subir archivo para Aprendices (.xlsx)", type=["xlsx"], key="up_apr_v2")
            
        if st.button("🧩 Procesar e Integrar Base de Datos Maestro", type="primary"):
            if file_cabezote is not None and file_aprendices is not None:
                try:
                    with st.spinner("Unificando el archivo maestro..."):
                        df_c = pd.read_excel(file_cabezote, header=None)
                        df_a = pd.read_excel(file_aprendices, header=None)
                        
                        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                            df_c.to_excel(writer, sheet_name="Cabezote", index=False, header=False)
                            df_a.to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                            
                    st.success(f"¡Excelente! El archivo maestro `{DB_FILE}` ha sido configurado correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Ocurrió un error al empaquetar el archivo: {e}")
            else:
                st.warning("Por favor cargue ambos archivos antes de procesar.")
