import streamlit as pd
import pandas as pd
import os

# Nombre del archivo de base de datos
DB_FILE = "Reporte de Asistencia.xlsx"

# ==========================================
# 0. FUNCIONES DE CARGA Y CONTROL DE DATOS
# ==========================================
def cargar_datos():
    """Carga de forma segura las hojas del archivo Excel."""
    try:
        df_cabezote = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
    except Exception:
        df_cabezote = pd.DataFrame()
        
    try:
        df_aprendices = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None)
    except Exception:
        df_aprendices = pd.DataFrame()
        
    try:
        df_instructores = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None)
    except Exception:
        df_instructores = pd.DataFrame()
        
    try:
        df_asistencias = pd.read_excel(DB_FILE, sheet_name="Asistencias", header=None)
    except Exception:
        df_asistencias = pd.DataFrame()
        
    try:
        df_notas = pd.read_excel(DB_FILE, sheet_name="Notas", header=None)
    except Exception:
        df_notas = pd.DataFrame()
        
    return df_cabezote, df_aprendices, df_instructores, df_asistencias, df_notas

# Inicializar datos
df_cabezote, df_aprendices, df_instructores, df_asistencias, df_notas = cargar_datos()

# ==========================================
# INTERFAZ DE STREAMLIT (ESTRUCTURA DE PESTAÑAS)
# ==========================================
st.title("📊 Sistema de Control de Asistencia y Notas - SENA")

# Creación de pestañas principales
tab1, tab2, tab3, tab4 = st.tabs([
    "🏠 Dashboard", 
    "📝 Llamado a Lista", 
    "🎯 Evaluar Competencia", 
    "⚙️ Configuración y Fichas"
])

# ------------------------------------------
# BARRA LATERAL (FILTROS GLOBALES)
# ------------------------------------------
st.sidebar.header("🔍 Filtros de Selección")

# Selector de Instructor
if not df_instructores.empty:
    instructores_lista = df_instructores.iloc[:, 0].dropna().unique().tolist()
    instructor_seleccionado = st.sidebar.selectbox("Seleccione Instructor:", instructores_lista)
else:
    instructor_seleccionado = st.sidebar.selectbox("Seleccione Instructor:", ["No hay instructores - Configure en Pestaña 4"])

# Filtrar Fichas/Grupos asignados a ese instructor en Cabezote
fichas_disponibles = []
if not df_cabezote.empty:
    # Columna 0: Instructor, Columna 1: Ficha/Grupo
    fichas_filtradas = df_cabezote[df_cabezote[0] == instructor_seleccionado]
    if not fichas_filtradas.empty:
        fichas_disponibles = fichas_filtradas[1].dropna().astype(str).unique().tolist()

if fichas_disponibles:
    ficha_seleccionada = st.sidebar.selectbox("Seleccione Ficha / Grupo:", fichas_disponibles)
else:
    ficha_seleccionada = st.sidebar.selectbox("Seleccione Ficha / Grupo:", ["Ninguna ficha asignada"])

# Conteo rápido de alumnos en la barra lateral
alumnos_grupo = pd.DataFrame()
if not df_aprendices.empty and ficha_seleccionada != "Ninguna ficha asignada":
    # Asumiendo Columna E (Índice 4) para Ficha/Grupo en el listado de aprendices
    alumnos_grupo = df_aprendices[df_aprendices[4].astype(str) == str(ficha_seleccionada)]
    st.sidebar.metric(label="Aprendices Activos", value=len(alumnos_grupo))


# ------------------------------------------
# PESTAÑA 1: DASHBOARD (Resumen de la Ficha)
# ------------------------------------------
with tab1:
    st.header("🏠 Resumen General")
    if ficha_seleccionada != "Ninguna ficha asignada":
        st.subheader(f"Información de la Ficha: {ficha_seleccionada}")
        # Buscar detalles del cabezote
        info_ficha = df_cabezote[(df_cabezote[0] == instructor_seleccionado) & (df_cabezote[1].astype(str) == str(ficha_seleccionada))]
        if not info_ficha.empty:
            st.info(f"**Trimestre:** {info_ficha.iloc[0, 2]} | **Competencia/Asignación:** {info_ficha.iloc[0, 3]}")
        
        st.markdown("### Listado Actual de Aprendices")
        if not alumnos_grupo.empty:
            # Mostrar columnas relevantes: L (Doc), M (Nombre 1), N (Apellido 1) -> Índices 11, 12, 13
            df_vista = alumnos_grupo.iloc[:, [11, 12, 13]].copy()
            df_vista.columns = ["Documento", "Nombre", "Apellido"]
            st.dataframe(df_vista, use_container_width=True)
        else:
            st.warning("No se encontraron aprendices asignados a esta ficha.")
    else:
        st.info("Por favor, seleccione o configure un instructor y una ficha en la barra lateral o en la pestaña 4.")


# ------------------------------------------
# PESTAÑA 2: LLAMADO A LISTA
# ------------------------------------------
with tab2:
    st.header("📝 Registro de Asistencia Diaria")
    if ficha_seleccionada == "Ninguna ficha asignada" or alumnos_grupo.empty:
        st.warning("⚠️ No hay aprendices activos asignados al grupo seleccionado. Registre la lista de aprendices o agregue la ficha en la Pestaña 4.")
    else:
        fecha_asistencia = st.date_input("Fecha de la sesión:", pd.Timestamp.now().date())
        st.markdown(f"**Tomando asistencia para la Ficha:** {ficha_seleccionada}")
        
        # Formulario simulado para asistencia
        asistencias_dict = {}
        for idx, row in alumnos_grupo.iterrows():
            doc = row[11]
            nombre_completo = f"{row[12]} {row[13]}"
            asistencias_dict[doc] = st.radio(f"{nombre_completo} ({doc})", ["Presente", "Falta Justificada", "Falta Injustificada"], key=f"asist_{doc}", horizontal=True)
            
        if st.button("💾 Guardar Asistencia del Día"):
            st.success(f"Asistencia del {fecha_asistencia} almacenada con éxito (Memoria Temporal).")


# ------------------------------------------
# PESTAÑA 3: EVALUAR COMPETENCIA
# ------------------------------------------
with tab3:
    st.header("🎯 Evaluación de Competencias (Juicios de Valor)")
    if ficha_seleccionada == "Ninguna ficha asignada" or alumnos_grupo.empty:
        st.warning("⚠️ No hay aprendices activos asignados al grupo seleccionado.")
    else:
        st.markdown(f"**Evaluación de Resultados de Aprendizaje para la Ficha:** {ficha_seleccionada}")
        resultado_evaluar = st.text_input("Resultado de Aprendizaje / Rap (Ej: RAP 1, Evidencia 1):")
        
        notas_dict = {}
        for idx, row in alumnos_grupo.iterrows():
            doc = row[11]
            nombre_completo = f"{row[12]} {row[13]}"
            notas_dict[doc] = st.selectbox(f"{nombre_completo} ({doc})", ["Aprobado (A)", "Deficiente (D)", "No Presentó"], key=f"nota_{doc}")
            
        if st.button("💾 Registrar Calificaciones"):
            st.success(f"Calificaciones para '{resultado_evaluar}' registradas con éxito (Memoria Temporal).")


# ------------------------------------------
# PESTAÑA 4: CONFIGURACIÓN Y GESTIÓN DE FICHAS
# ------------------------------------------
with tab4:
    st.header("⚙️ Configuración General de la Base de Datos")
    
    opcion_carga = st.radio(
        "Seleccione la acción a realizar:",
        ["✏️ Registrar Nueva Ficha / Modificar Cabezote", "📁 Subir Archivos Completos (.xlsx)"],
        horizontal=True
    )
    
    # OPCIÓN 1: FORMULARIO INTEGRADO (AQUÍ ESTÁ LA NUEVA LÓGICA)
    if opcion_carga == "✏️ Registrar Nueva Ficha / Modificar Cabezote":
        st.subheader("Formulario de Registro de Asignaciones")
        
        nuevo_grupo = st.text_input("Número de la Nueva Ficha / Grupo (Ej: 2831456):").strip()
        trimestre_seleccionado = st.selectbox("Trimestre del Año:", ["Trimestre I", "Trimestre II", "Trimestre III", "Trimestre IV"])
        nueva_asignacion = st.text_input("Nombre de la Competencia / Asignación:")
        
        if st.button("💾 Registrar Asignación / Actualizar Cabezote"):
            if nuevo_grupo:
                # 1. LEER LOS DATOS EXISTENTES
                df_cabezote_actual, df_aprendices_actual, df_instructores, df_asistencias, df_notas = cargar_datos()

                # 2. PREPARAR LA NUEVA FILA PARA EL CABEZOTE
                nueva_fila_cabezote = pd.DataFrame([[
                    instructor_seleccionado, 
                    nuevo_grupo, 
                    trimestre_seleccionado, 
                    nueva_asignacion
                ]])
                df_cabezote_final = pd.concat([df_cabezote_actual, nueva_fila_cabezote], ignore_index=True)

                # ==========================================
                # 🔥 CONTROL DE FICHA NUEVA AUTOMÁTICA
                # ==========================================
                ficha_existe = False
                if not df_aprendices_actual.empty and df_aprendices_actual.shape[1] > 4:
                    # Buscamos en la columna E (índice 4), saltándonos posibles encabezados
                    fichas_registradas = df_aprendices_actual.iloc[:, 4].dropna().astype(str).unique()
                    if str(nuevo_grupo) in fichas_registradas:
                        ficha_existe = True

                # Si la ficha NO existe en los aprendices, creamos el alumno de emergencia
                if not ficha_existe:
                    num_columnas = df_aprendices_actual.shape[1] if not df_aprendices_actual.empty else 16
                    fila_emergencia = [""] * num_columnas
                    
                    # Índices estándar SENA: E(4)=Ficha, L(11)=Doc, M(12)=Nombre, N(13)=Apellido
                    if num_columnas > 4:
                        fila_emergencia[4] = nuevo_grupo       # Columna E
                    if num_columnas > 11:
                        fila_emergencia[11] = 0                # Columna L (Documento provisional)
                    if num_columnas > 12:
                        fila_emergencia[12] = "NUEVA FICHA"     # Columna M
                    if num_columnas > 13:
                        fila_emergencia[13] = "(Sin alumnos)"   # Columna N
                    
                    nueva_fila_aprendiz = pd.DataFrame([fila_emergencia])
                    df_aprendices_final = pd.concat([df_aprendices_actual, nueva_fila_aprendiz], ignore_index=True)
                else:
                    df_aprendices_final = df_aprendices_actual

                # Asegurar que la hoja de instructores no esté vacía
                if df_instructores.empty:
                    df_instructores = pd.DataFrame([[instructor_seleccionado, "SENA2026"]])

                # 3. ESCRIBIR EN EL EXCEL REAL EN EL SERVIDOR
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    df_cabezote_final.to_excel(writer, sheet_name="Cabezote", index=False, header=False)
                    df_aprendices_final.to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                    df_instructores.to_excel(writer, sheet_name="Listado de instructores", index=False, header=False)
                    if not df_asistencias.empty:
                        df_asistencias.to_excel(writer, sheet_name="Asistencias", index=False, header=False)
                    if not df_notas.empty:
                        df_notas.to_excel(writer, sheet_name="Notas", index=False, header=False)

                st.success(f"¡Ficha {nuevo_grupo} registrada exitosamente en Cabezote y habilitada de inmediato!")
                st.rerun()
            else:
                st.error("Por favor, introduce un número de ficha válido.")

    # OPCIÓN 2: SUBIR ARCHIVOS INDEPENDIENTES
    elif opcion_carga == "📁 Subir Archivos Completos (.xlsx)":
        st.markdown("Sube las hojas de cálculo por separado para realizar un empaquetado general.")
        file_cabezote = st.file_uploader("Subir archivo para Cabezote (.xlsx)", type=["xlsx"])
        file_aprendices = st.file_uploader("Subir archivo para Aprendices (.xlsx)", type=["xlsx"])
        
        if st.button("🧩 Procesar e Integrar Base de Datos Maestro"):
            if file_cabezote and file_aprendices:
                with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                    pd.read_excel(file_cabezote, header=None).to_excel(writer, sheet_name="Cabezote", index=False, header=False)
                    pd.read_excel(file_aprendices, header=None).to_excel(writer, sheet_name="Listado de aprendices", index=False, header=False)
                    
                    # Conservar instructores
                    if os.path.exists(DB_FILE):
                        try:
                            pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None).to_excel(writer, sheet_name="Listado de instructores", index=False, header=False)
                        except Exception:
                            pd.DataFrame([[instructor_seleccionado, "SENA2026"]]).to_excel(writer, sheet_name="Listado de instructores", index=False, header=False)
                st.success("Configurado e integrado con éxito.")
                st.rerun()

    # ==========================================
    # 🔥 SECCIÓN INTEGRAL: BOTÓN DE DESCARGA REAL
    # ==========================================
    st.markdown("---")
    st.subheader("📥 Descargar Base de Datos Actualizada")
    st.markdown("Presiona este botón para bajarte el archivo Excel vivo del servidor con todas las fichas, instructores y cabezotes modificados.")
    
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as file:
            st.download_button(
                label="📥 Descargar Reporte de Asistencia.xlsx",
                data=file,
                file_name="Reporte de Asistencia_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.caption("Aún no se ha generado un archivo base en el servidor.")
