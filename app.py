import requests
import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime
from github import Github
from openpyxl import load_workbook

# Configuración de la página para entorno móvil y de escritorio
st.set_page_config(page_title="Control de Asistencia y Evaluación - SENA", layout="wide")

# Nombre del archivo original de Excel en tu repositorio de GitHub
DB_FILE = "Reporte de Asistencia.xlsx"
REPO_NAME = "semaforojo-web/Asistencia-SENA-CMM"

# ==========================================
# FUNCIÓN AUXILIAR: DESCARGAR EXCEL DE GITHUB A MEMORIA
# ==========================================
def descargar_excel_desde_github():
    """Descarga el archivo Excel binario directamente usando la URL raw de GitHub para evitar corrupciones."""
    try:
        # Construimos la URL Raw directa de tu archivo en GitHub
        # Reemplaza 'semaforojo-web/asistencia-sena-cmm' si tu usuario o repositorio cambiaron
        url_raw = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{DB_FILE}"
        
        # Si tu repositorio es PRIVADO, necesitamos enviar el token de seguridad en las cabeceras
        headers = {}
        if "GITHUB_TOKEN" in st.secrets:
            headers = {"Authorization": f"token {st.secrets['GITHUB_TOKEN']}"}
            
        response = requests.get(url_raw, headers=headers)
        
        if response.status_code == 200:
            # Retornamos los bytes puros y limpios del archivo de Excel
            return io.BytesIO(response.content)
        else:
            st.sidebar.error(f"⚠️ Error al acceder al archivo en GitHub (Código {response.status_code})")
    except Exception as e:
        st.sidebar.error(f"⚠️ Fallo en la descarga directa: {e}")
    return None
# ==========================================
# FUNCIÓN DE LÓGICA PERSISTENTE EN GITHUB (BLINDADA)
# ==========================================
def guardar_y_sincronizar_a_github(df_cabezote_final, df_aprendices_final, df_instructores, df_asistencias=pd.DataFrame(), df_notas=pd.DataFrame()):
    """Guarda los datos en el Excel forzando la apertura nativa de openpyxl para que NUNCA elimine otras hojas."""
    if "GITHUB_TOKEN" not in st.secrets:
        st.warning("⚠️ No se detectó el 'GITHUB_TOKEN' en los Secrets de Streamlit.")
        return

    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(REPO_NAME)
        
        archivo_base = descargar_excel_desde_github()
        
        try:
            contents = repo.get_contents(DB_FILE, ref="main")
            sha = contents.sha
            wb = load_workbook(archivo_base)
        except Exception:
            from openpyxl import Workbook
            wb = Workbook()
            sha = None

        def escribir_dataframe_en_hoja(wb_objeto, df_datos, nombre_hoja):
            if df_datos.empty:
                return
            if nombre_hoja in wb_objeto.sheetnames:
                ws = wb_objeto[nombre_hoja]
                ws.delete_rows(1, ws.max_row + 1)
                ws.delete_cols(1, ws.max_column + 1)
            else:
                ws = wb_objeto.create_sheet(title=nombre_hoja)
            
            for r_idx, row in enumerate(df_datos.values, start=1):
                for c_idx, value in enumerate(row, start=1):
                    if pd.isna(value):
                        ws.cell(row=r_idx, column=c_idx, value="")
                    else:
                        ws.cell(row=r_idx, column=c_idx, value=value)

        # 2. Actualizamos de forma aislada únicamente nuestras hojas de la app
        escribir_dataframe_en_hoja(wb, df_cabezote_final, "Cabezote")
        escribir_dataframe_en_hoja(wb, df_aprendices_final, "Listado de aprendices")
        escribir_dataframe_en_hoja(wb, df_instructores, "Listado de instructores")
        if not df_asistencias.empty:
            escribir_dataframe_en_hoja(wb, df_asistencias, "Asistencias")
        if not df_notas.empty:
            escribir_dataframe_en_hoja(wb, df_notas, "Notas")

        # ========================================================
        # INYECCIÓN AUTOMÁTICA DE LA FÓRMULA EN LA HOJA CABEZOTE
        # ========================================================
        ws_cabezote = wb["Cabezote"]
        
        # Opción B (Dinámica): Se aplica automáticamente a la ÚLTIMA fila real escrita
        ultima_fila = ws_cabezote.max_row
        if ultima_fila >= 1:
            ws_cabezote[f"A{ultima_fila}"] = f"=VLOOKUP(G{ultima_fila},'Listado de aprendices'!$E:$I,5,0)"
        # ========================================================

        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            wb.remove(wb["Sheet"])

        output = io.BytesIO()
        wb.save(output)
        content = output.getvalue()
        
        try:
            contents_verificar = repo.get_contents(DB_FILE, ref="main")
            sha = contents_verificar.sha
        except Exception:
            sha = None

        if sha:
            repo.update_file(
                path=DB_FILE,
                message="🤖 Actualización nativa de celdas (Preservando el 100% de hojas originales)",
                content=content,
                sha=sha,
                branch="main"
            )
        else:
            repo.create_file(
                path=DB_FILE,
                message="🤖 Creación inicial segura del archivo maestro",
                content=content,
                branch="main"
            )
        st.success("🔄 ¡Datos sincronizados! El resto de tus hojas originales han sido preservadas intactas.")
    except Exception as e:
        st.error(f"⚠️ Error crítico en la conexión con GitHub: {e}")

        
# ==========================================
# FUNCIONES DE LECTURA DIRECTA DESDE GITHUB
# ==========================================
def obtener_instructores_y_contraseñas():
    """Lee los instructores y sus contraseñas desde GitHub de forma segura forzando el motor openpyxl"""
    dict_usuarios = {}
    lista_instructores = []
    
    archivo_memoria = descargar_excel_desde_github()
    if archivo_memoria:
        try:
            # Añadimos engine='openpyxl' para garantizar la lectura de bytes binarios
            df_inst = pd.read_excel(archivo_memoria, sheet_name="Listado de instructores", header=None, engine='openpyxl')
            for idx, row in df_inst.iterrows():
                nombre = str(row[0]).strip() if pd.notna(row[0]) else ""
                password = str(row[1]).strip() if df_inst.shape[1] > 1 and pd.notna(row[1]) else "SENA2026"
                
                if nombre.upper() != "INSTRUCTOR" and nombre != "" and nombre.upper() != "NAN":
                    dict_usuarios[nombre] = password
                    lista_instructores.append(nombre)
                    
            if lista_instructores:
                return sorted(lista_instructores), dict_usuarios
            else:
                return ["La hoja está vacía"], {}
        except Exception as e:
            return [f"Error al leer hoja: {e}"], {}
            
    return ["No se pudo conectar a GitHub"], {}
def cargar_datos():
    """Carga y procesa el listado de aprendices desde GitHub asegurando el motor openpyxl"""
    archivo_memoria = descargar_excel_desde_github()
    if archivo_memoria:
        try:
            # Forzamos el uso de engine='openpyxl' para evitar errores de lectura binaria
            df = pd.read_excel(archivo_memoria, sheet_name="Listado de aprendices", header=None, engine='openpyxl')
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
            st.sidebar.error(f"⚠️ Error al filtrar listado de aprendices: {e}")
            
    return pd.DataFrame(columns=["Grupo", "Documento", "Nombre Completo"])

def obtener_trimestres_disponibles(grupo, instructor):
    """Busca los trimestres vinculados a un grupo e instructor en el Cabezote desde GitHub"""
    archivo_memoria = descargar_excel_desde_github()
    if archivo_memoria:
        try:
            df_cab = pd.read_excel(archivo_memoria, sheet_name="Cabezote", header=None)
            df_cab[5] = df_cab[5].astype(str).str.strip()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            
            filtro = (df_cab[6] == str(grupo)) & (df_cab[5].str.upper() == str(instructor).strip().upper())
            resultado = df_cab[filtro]
            
            if not resultado.empty and resultado.shape[1] > 48:
                trimestres = resultado.iloc[:, 48].dropna().astype(str).str.strip().unique().tolist()
                trimestres_validos = sorted([t for t in trimestres if t != "" and t.upper() != "NAN" and t.upper() != "TRIMESTRE"])
                if trimestres_validos:
                    return trimestres_validos
        except Exception:
            pass
    return ["Sin trimestres detectados"]

def obtener_materias_disponibles(grupo, instructor, trimestre):
    """Obtiene los números de asignación cargados desde GitHub"""
    archivo_memoria = descargar_excel_desde_github()
    if archivo_memoria:
        try:
            df_cab = pd.read_excel(archivo_memoria, sheet_name="Cabezote", header=None)
            df_cab[5] = df_cab[5].astype(str).str.strip()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            df_cab[47] = df_cab[47].astype(str).str.strip() if df_cab.shape[1] > 47 else ""
            
            filtro = (df_cab[6] == str(grupo)) & \
                     (df_cab[5].str.upper() == str(instructor).strip().upper()) & \
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
    """Retorna el nombre largo de la asignatura basándose en el archivo de GitHub"""
    archivo_memoria = descargar_excel_desde_github()
    if archivo_memoria:
        try:
            df_cab = pd.read_excel(archivo_memoria, sheet_name="Cabezote", header=None)
            df_cab[3] = df_cab[3].astype(str).str.strip()
            df_cab[5] = df_cab[5].astype(str).str.strip()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            df_cab[47] = df_cab[47].astype(str).str.strip() if df_cab.shape[1] > 47 else ""
            
            filtro = (df_cab[6] == str(grupo)) & \
                     (df_cab[5].str.upper() == str(instructor).strip().upper()) & \
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
    return "Archivo Excel no encontrado en GitHub"
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
                    st.error("Por favor, cargue la estructura de instructores correcta en su Excel de GitHub o configure sus Secrets.")
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

lista_grupos = sorted(df_aprendices["Grupo"].dropna().unique()) if not df_aprendices.empty else []
grupo_seleccionado = st.sidebar.selectbox("1. Seleccione el Grupo:", lista_grupos if lista_grupos else ["Sin datos"])

lista_trimestres_dinamicos = obtener_trimestres_disponibles(grupo_seleccionado, instructor_seleccionado)
trimestre_seleccionado = st.sidebar.selectbox("3. Seleccione el Trimestre:", lista_trimestres_dinamicos)

lista_asignaciones_dinamicas = obtener_materias_disponibles(grupo_seleccionado, instructor_seleccionado, trimestre_seleccionado)
asignacion_num_seleccionada = st.sidebar.selectbox("4. Seleccione Asignación:", lista_asignaciones_dinamicas)

materia_detectada = filtrar_materia_final(grupo_seleccionado, instructor_seleccionado, trimestre_seleccionado, asignacion_num_seleccionada)

st.sidebar.markdown("---")
st.sidebar.info(f"📚 **Materia Vinculada:**\n{materia_detectada}")

alumnos_grupo = df_aprendices[df_aprendices["Grupo"] == grupo_seleccionado].reset_index(drop=True) if not df_aprendices.empty else pd.DataFrame()
st.sidebar.markdown(f"**Total Aprendices Activos:** {len(alumnos_grupo)}")

tab1, tab2, tab3, tab4 = st.tabs(["📋 Llamado a Lista", "📝 Evaluar Competencia", "📈 Historial y Reportes", "📂 Alimentar y Cargar Bases"])

# PESTAÑA 1: LLAMADO A LISTA
with tab1:
    st.header(f"📋 Control de Asistencia")
    fecha_asistencia = st.date_input("Fecha del llamado a lista:", datetime.now())
    
    if alumnos_grupo.empty:
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
            
            df_nuevos_pasos = pd.DataFrame(registros)
            df_nuevos_pasos.to_csv("asistencia_guardada.csv", mode='a', header=not os.path.exists("asistencia_guardada.csv"), index=False)
            
            if "GITHUB_TOKEN" in st.secrets:
                try:
                    g = Github(st.secrets["GITHUB_TOKEN"])
                    repo = g.get_repo(REPO_NAME)
                    with open("asistencia_guardada.csv", "r", encoding='utf-8') as f:
                        contenido_csv = f.read()
                    try:
                        sha = repo.get_contents("asistencia_guardada.csv", ref="main").sha
                        repo.update_file("asistencia_guardada.csv", "🤖 Actualizar histórico asistencias CSV", contenido_csv, sha, branch="main")
                    except Exception:
                        repo.create_file("asistencia_guardada.csv", "🤖 Crear histórico asistencias CSV", contenido_csv, branch="main")
                    st.success("🔄 ¡Historial de Asistencias respaldado en GitHub!")
                except Exception as e:
                    st.warning(f"Guardado local, pero no se pudo subir a GitHub: {e}")

# PESTAÑA 2: EVALUACIÓN
with tab2:
    st.header(f"📝 Registro de Juicios Evaluativos")
    fecha_evaluacion = st.date_input("Fecha de Evaluación:", datetime.now())
    
    if alumnos_grupo.empty:
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
            
            df_nuevas_notas = pd.DataFrame(registros_eval)
            df_nuevas_notas.to_csv("evaluaciones_guardadas.csv", mode='a', header=not os.path.exists("evaluaciones_guardadas.csv"), index=False)
            
            if "GITHUB_TOKEN" in st.secrets:
                try:
                    g = Github(st.secrets["GITHUB_TOKEN"])
                    repo = g.get_repo(REPO_NAME)
                    with open("evaluaciones_guardadas.csv", "r", encoding='utf-8') as f:
                        contenido_notas = f.read()
                    try:
                        sha = repo.get_contents("evaluaciones_guardadas.csv", ref="main").sha
                        repo.update_file("evaluaciones_guardadas.csv", "🤖 Actualizar histórico notas CSV", contenido_notas, sha, branch="main")
                    except Exception:
                        repo.create_file("evaluaciones_guardadas.csv", "🤖 Crear histórico notas CSV", contenido_notas, branch="main")
                    st.success("🔄 ¡Historial de Notas respaldado en GitHub!")
                except Exception as e:
                    st.warning(f"Guardado local, pero no se pudo subir a GitHub: {e}")

# PESTAÑA 3: REPORTES
with tab3:
    st.header("📈 Historial de Registros")
    sub_tab1, sub_tab2 = st.tabs(["Histórico de Asistencias", "Histórico de Notas"])
    
    with sub_tab1:
        if os.path.exists("asistencia_guardada.csv"):
            df_asist_hist = pd.read_csv("asistencia_guardada.csv")
            st.dataframe(df_asist_hist, use_container_width=True)
            
            csv_asist = df_asist_hist.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Descargar Historial de Asistencias Completo (CSV)",
                data=csv_asist,
                file_name="asistencia_guardada_completa.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("No hay registros de asistencia locales aún.")
            
    with sub_tab2:
        if os.path.exists("evaluaciones_guardadas.csv"):
            df_eval_hist = pd.read_csv("evaluaciones_guardadas.csv")
            st.dataframe(df_eval_hist, use_container_width=True)
            
            csv_notas = df_eval_hist.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Descargar Historial de Notas Completo (CSV)",
                data=csv_notas,
                file_name="evaluaciones_guardadas_completa.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("No hay registros de notas locales aún.")

# PESTAÑA 4: GESTIÓN DE BASES DE DATOS
with tab4:
    st.header("📂 Gestión y Alimentación de la Base de Datos")
    opcion_carga = st.radio("Seleccione el método para gestionar datos:", ["✍️ Alimentar Cabezote Directamente (Formulario)", "📁 Subir Archivos Completos (.xlsx)"])
    
    if opcion_carga == "✍️ Alimentar Cabezote Directamente (Formulario)":
        with st.form("form_registro_directo_cabezote_completo"):
            c1, c2, c3, c4 = st.columns(4)
            input_grupo = c1.text_input("Número de Grupo (Columna G):", placeholder="Ej: 3141501")
            input_instructor = c2.text_input("Nombre del Instructor (Columna F):", value=instructor_seleccionado, disabled=True)
            input_asignacion_num = c3.selectbox("Número de asignación (Columna D):", ["1", "2", "3"])
            input_resultados = c4.text_input("Resultados de Aprendizaje (Columna K):", placeholder="Ej: Mantenimiento")
            
            st.markdown("##### 🚀 Planificación Estratégica del Proyecto")
            c5, c6, c7 = st.columns(3)
            input_fase = c5.text_input("Fase del proyecto (Columna H):")
            input_actividades = c6.text_area("Actividades del Proyecto (Columna I):", height=68)
            input_competencia = c7.text_area("Competencia (Columna J):", height=68)
            
            c8, c9, c10, c11 = st.columns([3, 1, 1, 1])
            input_li = c8.text_input("Linea (Columna L):")
            input_tri = c9.text_input("Trimestre rep (Columna O):")
            input_horas = c10.text_input("Número de Horas (Columna M):")
            input_fecha_ini = c11.text_input("Fecha de inicio (Columna N):", placeholder="DD/MM/AAAA")
            
            st.markdown("##### 🏫 Datos Locales del Ambiente y Jornada")
            c12, c13, c14, c15 = st.columns(4)
            input_ambiente = c12.text_input("Ambiente (Columna P):")
            input_dia = c13.text_input("Día (Columna Q):")
            input_horario = c14.text_input("Horario (Columna R):")
            input_jornada = c15.text_input("Jornada (Columna S):")
            
            st.markdown("##### 📑 Evidencias del Proceso (1 al 5)")
            col_ev1, col_ev2, col_ev3, col_ev4, col_ev5 = st.columns(5)
            input_ev1 = col_ev1.text_input("Evidencia 1:")
            input_ev2 = col_ev2.text_input("Evidencia 2:")
            input_ev3 = col_ev3.text_input("Evidencia 3:")
            input_ev4 = col_ev4.text_input("Evidencia 4:")
            input_ev5 = col_ev5.text_input("Evidencia 5:")
            
            with st.expander("🛠️ Descripciones Prácticas (1 al 11)"):
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
                input_trimestre = c_f1.text_input("Trimestre en curso (Columna AV):", placeholder="Ej: Trimestre 1")
                input_observaciones = c_f2.text_input("Observaciones (Columna AW):")
            
            boton_agregar_cab = st.form_submit_button("💾 Insertar y Sincronizar en GitHub", type="primary")
            
        if boton_agregar_cab:
            if input_grupo and input_instructor and input_trimestre:
                try:
                    df_cab_existente = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    df_apr_existente = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    df_inst_existente = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None) if os.path.exists(DB_FILE) else pd.DataFrame()
                    
                    ancho_columnas = 49
                    nueva_fila = [""] * ancho_columnas
                    if str(input_asignacion_num).strip().isdigit():
                       nueva_fila[3] = int(input_asignacion_num)
                    else:
                       nueva_fila[3] = str(input_asignacion_num).strip()
            
                    nueva_fila[5] = str(input_instructor).strip()
                    grupo_limpio = str(input_grupo).strip()
                    if grupo_limpio.isdigit():
                       nueva_fila[6] = int(grupo_limpio)
                    else:
                       nueva_fila[6] = grupo_limpio
                    nueva_fila[6] = str(input_grupo).strip()
                    nueva_fila[7] = str(input_fase).strip()
                    nueva_fila[8] = str(input_actividades).strip()
                    nueva_fila[9] = str(input_competencia).strip()
                    nueva_fila[10] = str(input_resultados).strip()
                    nueva_fila[11] = str(input_li).strip()
                    nueva_fila[12] = str(input_horas).strip()
                    nueva_fila[13] = str(input_fecha_ini).strip()
                    nueva_fila[14] = str(input_tri).strip()
                    nueva_fila[15] = str(input_ambiente).strip()
                    nueva_fila[16] = str(input_dia).strip()
                    nueva_fila[17] = str(input_horario).strip()
                    nueva_fila[18] = str(input_jornada).strip()
                    
                    nueva_fila[19] = str(input_ev1).strip()
                    nueva_fila[20] = str(input_ev2).strip()
                    nueva_fila[21] = str(input_ev3).strip()
                    nueva_fila[22] = str(input_ev4).strip()
                    nueva_fila[23] = str(input_ev5).strip()
                    
                    nueva_fila[24] = str(input_p1).strip()
                    nueva_fila[25] = str(input_p2).strip()
                    nueva_fila[26] = str(input_p3).strip()
                    nueva_fila[27] = str(input_p4).strip()
                    nueva_fila[28] = str(input_p5).strip()
                    nueva_fila[29] = str(input_p6).strip()
                    nueva_fila[30] = str(input_p7).strip()
                    nueva_fila[31] = str(input_p8).strip()
                    nueva_fila[32] = str(input_p9).strip()
                    nueva_fila[33] = str(input_p10).strip()
                    nueva_fila[34] = str(input_p11).strip()
                    
                    nueva_fila[35] = str(input_s1).strip()
                    nueva_fila[36] = str(input_s2).strip()
                    nueva_fila[37] = str(input_s3).strip()
                    nueva_fila[38] = str(input_s4).strip()
                    nueva_fila[39] = str(input_s5).strip()
                    nueva_fila[40] = str(input_s6).strip()
                    nueva_fila[41] = str(input_s7).strip()
                    nueva_fila[42] = str(input_s8).strip()
                    nueva_fila[43] = str(input_s9).strip()
                    nueva_fila[44] = str(input_s10).strip()
                    nueva_fila[45] = str(input_s11).strip()
                    
                    nueva_fila[46] = str(input_observaciones).strip()
                    nueva_fila[47] = str(input_trimestre).strip()
                    
                    df_cab_final = pd.concat([df_cab_existente, pd.DataFrame([nueva_fila])], ignore_index=True)
                    
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
                    "Número Asignación": df_preview[3], "Resultados": df_preview[10], "Trimestre": df_preview[47]
                }).dropna(subset=["Grupo", "Instructor"], how="all")
                st.dataframe(df_resumen, use_container_width=True)
                
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
