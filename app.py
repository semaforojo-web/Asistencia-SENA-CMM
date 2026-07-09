import streamlit as st
import pandas as pd
import os
from datetime import datetime
from github import Github

# Configuración de la página para entorno móvil y de escritorio
st.set_page_config(page_title="Control de Asistencia y Evaluación - SENA", layout="wide")

# Nombre del archivo original de Excel en tu repositorio de GitHub
DB_FILE = "Reporte de Asistencia.xlsx"
REPO_NAME = "semaforojo-web/asistencia-sena-cmm"  # 👈 CAMBIA ESTO por tu usuario y repositorio real

# ==========================================
# FUNCIÓN DE LOGICA PERSISTENTE EN GITHUB
# ==========================================
def guardar_y_sincronizar_a_github(df_cabezote_final, df_aprendices_final, df_instructores, df_asistencias=pd.DataFrame(), df_notas=pd.DataFrame()):
    """Guarda el Excel localmente y lo sube automáticamente al repositorio de GitHub."""
    # 1. Guardar localmente en el servidor temporal
    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
        df_cabezote_final.to_excel(writer, sheet_name="Cabezote", index=False, header=False)
        df_aprendices_final.to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
        df_instructores.to_excel(writer, sheet_name="Listado de instructores", index=False, header=False)
        if not df_asistencias.empty:
            df_asistencias.to_excel(writer, sheet_name="Asistencias", index=False, header=False)
        if not df_notas.empty:
            df_notas.to_excel(writer, sheet_name="Notas", index=False, header=False)
            
    # 2. Sincronizar automáticamente con GitHub via API
    if "GITHUB_TOKEN" in st.secrets:
        try:
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(REPO_NAME)
            
            with open(DB_FILE, "rb") as f:
                content = f.read()
                
            # Conseguir el SHA del archivo existente en GitHub para poder reemplazarlo
            try:
                contents = repo.get_contents(DB_FILE, ref="main")
                sha = contents.sha
            except Exception:
                sha = None # Por si el archivo no existe en el repositorio aún
            
            if sha:
                repo.update_file(
                    path=DB_FILE,
                    message="🤖 Actualización automática desde App Streamlit [Cabezote/Asistencia]",
                    content=content,
                    sha=sha,
                    branch="main"
                )
            else:
                repo.create_file(
                    path=DB_FILE,
                    message="🤖 Creación automática desde App Streamlit [Cabezote/Asistencia]",
                    content=content,
                    branch="main"
                )
            st.success("🔄 ¡Datos sincronizados con tu repositorio de GitHub con éxito!")
        except Exception as e:
            st.error(f"⚠️ El archivo se guardó localmente pero falló la sincronización con GitHub: {e}")
    else:
        st.warning("⚠️ No se detectó el 'GITHUB_TOKEN' en los Secrets de Streamlit.")

def obtener_instructores_y_contraseñas():
    """Lee los instructores (Columna A) y sus contraseñas (Columna B)"""
    dict_usuarios = {}
    lista_instructores = []
    
    if os.path.exists(DB_FILE):
        try:
            df_inst = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None)
            for idx, row in df_inst.iterrows():
                nombre = str(row[0]).strip() if pd.notna(row[0]) else ""
                password = str(row[1]).strip() if df_inst.shape[1] > 1 and pd.notna(row[1]) else "SENA2026"
                
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
            df_procesado["Grupo"] = df.iloc[:, 4].astype(str).str.strip()
            df_procesado["Documento"] = df.iloc[:, 11].fillna("S/D").astype(str).str.strip()
            
            nombres_completos = df.iloc[:, 12:15].fillna("").astype(str)
            df_procesado["Nombre Completo"] = nombres_completos.agg(' '.join, axis=1).str.replace(r'\s+', ' ', regex=True).str.strip()
            
            df_procesado["Estado"] = df.iloc[:, 15].fillna("").astype(str).str.upper().str.strip()
            terminos_excluir = "CANCELADO|RETIRO VOLUNTARIO|TRASLADO"
            df_procesado = df_procesado[~df_procesado["Estado"].str.contains(terminos_excluir, regex=True)]
            
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
    """Retorna el nombre largo de la asignatura basándose en el número de asignación"""
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
                        st.error("Contraseña incorrecta. Inténtelo de nuevo.")
                else:
                    st.error("Por favor, cargue la estructura de instructores correcta en su Excel de GitHub.")
    st.stop()

# --- EJECUCIÓN PRINCIPAL ---
df_aprendices = cargar_datos()
instructor_seleccionado = st.session_state["instructor_logueado"]

if not os.path.exists("asistencia_guardada.csv"):
    pd.DataFrame(columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Asignacion_Num", "Resultados", "Documento", "Nombre", "Asistencia"]).to_csv("asistencia_guardada.csv", index=False)

if not os.path.exists("evaluaciones_guardadas.csv"):
    pd.DataFrame(columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Materia", "Documento", "Nombre", "Evaluación (A/D)", "Observaciones"]).to_csv("evaluaciones_guardadas.csv", index=False)

st.title("📊 Dashboard de Gestión de Ambientes de Formación - SENA")

if st.sidebar.button("🔒 Cerrar Sesión"):
    st.session_state["autenticado"] = False
    st.session_state["instructor_logueado"] = None
    st.rerun()

st.sidebar.markdown(f"👤 **Instructor activo:**\n`{instructor_seleccionado}`")
st.sidebar.markdown("---")
st.sidebar.header("⚙️ Filtros de Planificación")

lista_grupos = sorted(df_aprendices["Grupo"].dropna().unique())
grupo_seleccionado = st.sidebar.selectbox("1. Seleccione el Grupo:", lista_grupos if lista_grupos else ["Error"])

lista_trimestres_dinamicos = obtener_trimestres_disponibles(grupo_seleccionado, instructor_seleccionado)
trimestre_seleccionado = st.sidebar.selectbox("3. Seleccione el Trimestre:", lista_trimestres_dinamicos)

lista_asignaciones_dinamicas = obtener_materias_disponibles(grupo_seleccionado, instructor_seleccionado, trimestre_seleccionado)
asignacion_num_seleccionada = st.sidebar.selectbox("4. Seleccione Asignación:", lista_asignaciones_dinamicas)

materia_detectada = filtrar_materia_final(grupo_seleccionado, instructor_seleccionado, trimestre_seleccionado, asignacion_num_seleccionada)

st.sidebar.markdown("---")
st.sidebar.info(f"📚 **Materia Vinculada:**\n{materia_detectada}")

alumnos_grupo = df_aprendices[df_aprendices["Grupo"] == grupo_seleccionado].reset_index(drop=True)
st.sidebar.markdown(f"**Total Aprendices Activos:** {len(alumnos_grupo)}")

tab1, tab2, tab3, tab4 = st.tabs(["📋 Llamado a Lista", "📝 Evaluar Competencia", "📈 Historial y Reportes", "📂 Alimentar y Cargar Bases"])

# PESTAÑA 1: LLAMADO A LISTA
with tab1:
    st.header(f"📋 Control de Asistencia")
    fecha_asistencia = st.date_input("Fecha del llamado a lista:", datetime.now())
    
    if alumnos_grupo.empty or grupo_seleccionado == "Error":
        st.warning(f"No hay aprendices activos asignados al grupo seleccionado.")
    else:
        with st.form(key=f"formulario_asistencia_{grupo_seleccionado}"):
            asistencia_dict = {}
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
                registros.append({
                    "Fecha": fecha_asistencia, "Grupo": grupo_seleccionado, "Instructor": instructor_seleccionado,
                    "Trimestre": trimestre_seleccionado, "Asignacion_Num": asignacion_num_seleccionada,
                    "Resultados": materia_detectada, "Documento": row["Documento"], "Nombre": row["Nombre Completo"], "Asistencia": estado
                })
            pd.DataFrame(registros).to_csv("asistencia_guardada.csv", mode='a', header=False, index=False)
            st.success("¡Asistencia guardada localmente!")

# PESTAÑA 2: EVALUACIÓN
with tab2:
    st.header(f"📝 Registro de Juicios Evaluativos")
    fecha_evaluacion = st.date_input("Fecha de Evaluación:", datetime.now())
    
    if alumnos_grupo.empty or grupo_seleccionado == "Error":
        st.warning("No hay alumnos para evaluar.")
    else:
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
                registros_eval.append({
                    "Fecha": fecha_evaluacion, "Grupo": grupo_seleccionado, "Instructor": instructor_seleccionado,
                    "Trimestre": trimestre_seleccionado, "Materia": materia_detectada, "Documento": row["Documento"],
                    "Nombre": row["Nombre Completo"], "Evaluación (A/D)": eval_dict[idx], "Observaciones": obs_dict[idx]
                })
            pd.DataFrame(registros_eval).to_csv("evaluaciones_guardadas.csv", mode='a', header=False, index=False)
            st.success("Evaluaciones guardadas.")

# PESTAÑA 3: REPORTES
with tab3:
    st.header("📈 Historial de Registros")
    sub_tab1, sub_tab2 = st.tabs(["Histórico de Asistencias", "Histórico de Notas"])
    with sub_tab1:
        if os.path.exists("asistencia_guardada.csv"):
            df_asist_hist = pd.read_csv("asistencia_guardada.csv")
            st.dataframe(df_asist_hist, use_container_width=True)
    with sub_tab2:
        if os.path.exists("evaluaciones_guardadas.csv"):
            df_eval_hist = pd.read_csv("evaluaciones_guardadas.csv")
            st.dataframe(df_eval_hist, use_container_width=True)

# PESTAÑA 4: GESTIÓN DE BASES DE DATOS (CON CONEXIÓN DIRECTA A GITHUB)
with tab4:
    st.header("📂 Gestión y Alimentación de la Base de Datos")
    opcion_carga = st.radio("Seleccione el método para gestionar datos:", ["✍️ Alimentar Cabezote Directamente (Formulario)", "📁 Subir Archivos Completos (.xlsx)"])
    
    if opcion_carga == "✍️ Alimentar Cabezote Directamente (Formulario)":
        with st.form("form_registro_directo_cabezote"):
            c1, c2 = st.columns(2)
            input_grupo = c1.text_input("Número de Grupo:")
            input_instructor = c2.text_input("Nombre del Instructor:", value=instructor_seleccionado, disabled=True)
            c3, c4 = st.columns(2)
            input_asignacion_num = c3.selectbox("Número de asignación:", ["1", "2", "3"])
            input_trimestre = c4.text_input("Trimestre de la formación:")
            input_resultados = st.text_input("Resulados de Aprendizaje:")
            boton_agregar_cab = st.form_submit_button("💾 Insertar y Sincronizar en GitHub", type="primary")
            
        if boton_agregar_cab:
            if input_grupo and input_instructor and input_resultados and input_trimestre:
                try:
                    df_cab_existente = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    df_apr_existente = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    df_inst_existente = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    
                    ancho_columnas = max(df_cab_existente.shape[1], 48) if not df_cab_existente.empty else 48
                    nueva_fila = [""] * ancho_columnas
                    nueva_fila[3] = str(input_asignacion_num).strip()
                    nueva_fila[5] = str(input_instructor).strip()
                    nueva_fila[6] = str(input_grupo).strip()
                    nueva_fila[10] = str(input_resultados).strip()
                    nueva_fila[47] = str(input_trimestre).strip()
                    
                    df_cab_final = pd.concat([df_cab_existente, pd.DataFrame([nueva_fila])], ignore_index=True)
                    
                    # Llamamos a la sincronización en lugar de guardar local únicamente
                    guardar_y_sincronizar_a_github(df_cab_final, df_apr_existente, df_inst_existente)
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al escribir en el archivo: {e}")

    elif opcion_carga == "📁 Subir Archivos Completos (.xlsx)":
        file_cabezote = st.file_uploader("Subir archivo para Cabezote (.xlsx)", type=["xlsx"])
        file_aprendices = st.file_uploader("Subir archivo para Aprendices (.xlsx)", type=["xlsx"])
        
        if st.button("🧩 Procesar e Integrar Base de Datos Maestro", type="primary"):
            if file_cabezote and file_aprendices:
                df_c = pd.read_excel(file_cabezote, header=None)
                df_a = pd.read_excel(file_aprendices, header=None)
                df_i = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None) if os.path.exists(DB_FILE) else pd.DataFrame([[instructor_seleccionado, "SENA2026"]])
                
                guardar_y_sincronizar_a_github(df_c, df_a, df_i)
                st.rerun()

    # --- VISTA PREVIA Y DESCARGA CSV ---
    if os.path.exists(DB_FILE):
        try:
            st.markdown("---")
            st.markdown("### 📋 Vista Previa del Cabezote Real:")
            df_preview = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            if df_preview.shape[1] > 47:
                df_resumen = pd.DataFrame({
                    "Grupo": df_preview[6], "Instructor": df_preview[5],
                    "Número Asignación": df_preview[3], "Asignatura": df_preview[10], "Trimestre": df_preview[47]
                }).dropna(subset=["Grupo", "Instructor"], how="all")
                st.dataframe(df_resumen, use_container_width=True)
                
                # Botón para descargar la vista previa en formato CSV
                csv_data = df_resumen.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Descargar esta vista previa en formato CSV",
                    data=csv_data,
                    file_name="vista_previa_cabezote.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        except Exception:
            pass
