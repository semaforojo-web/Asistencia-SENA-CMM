import os
import io
import requests
import holidays
import pandas as pd
import streamlit as st
from datetime import datetime
from github import Github
from openpyxl import load_workbook

# ==========================================
# CONFIGURACIÓN DE PÁGINA Y CONSTANTES
# ==========================================
st.set_page_config(page_title="Control de Asistencia y Evaluación - SENA", layout="wide")

DB_FILE = "Reporte de Asistencia.xlsx"
REPO_NAME = "semaforojo-web/Asistencia-SENA-CMM"

# Inicialización de archivos CSV locales si no existen
if not os.path.exists("asistencia_guardada.csv"):
    pd.DataFrame(
        columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Asignacion_Num", "Resultados", "Documento", "Nombre", "Asistencia"]
    ).to_csv("asistencia_guardada.csv", index=False)

if not os.path.exists("evaluaciones_guardadas.csv"):
    pd.DataFrame(
        columns=["Fecha", "Grupo", "Instructor", "Trimestre", "Materia", "Documento", "Nombre", "Evaluación (A/D)", "Observaciones"]
    ).to_csv("evaluaciones_guardadas.csv", index=False)


# ==========================================
# FUNCIÓN: CALCULAR FALLAS ACUMULADAS (FACU)
# ==========================================
def calcular_fallas_acumuladas(archivo_csv="asistencia_guardada.csv"):
    """
    Lee asistencia_guardada.csv, calcula el total de fallas ("IS") por aprendiz
    y aplica la lógica de días hábiles de Colombia para la columna FACU:
    - Muestra '2 C' cuando la 2.ª falla consecutiva cumple 5 días hábiles.
    - Muestra '2 C (X)' (ej. '2 C (3)') si la racha de 2C está madura y el aprendiz
      suma más fallas que aún no cumplen sus 5 días hábiles.
    - Muestra '3 C' cuando la 3.ª falla consecutiva cumple sus 5 días hábiles.
    """
    cols_base = ["Grupo", "Instructor", "Trimestre", "Asignacion_Num", "Documento", "Nombre", "FACU"]
    
    if not os.path.exists(archivo_csv):
        return pd.DataFrame(columns=cols_base)
    
    try:
        df_asistencia = pd.read_csv(archivo_csv)
        if df_asistencia.empty:
            return pd.DataFrame(columns=cols_base)
        
        cols_agrupacion = ["Grupo", "Instructor", "Trimestre", "Asignacion_Num", "Documento", "Nombre"]
        
        for col in cols_agrupacion + ["Asistencia", "Fecha"]:
            if col not in df_asistencia.columns:
                return pd.DataFrame(columns=cols_base)
        
        df_asistencia["Fecha"] = pd.to_datetime(df_asistencia["Fecha"], errors="coerce")
        df_asistencia = df_asistencia.sort_values(by=cols_agrupacion + ["Fecha"])
        
        fecha_actual = pd.Timestamp.now().normalize()
        
        # Festivos nacionales de Colombia
        anios = [fecha_actual.year - 1, fecha_actual.year, fecha_actual.year + 1]
        festivos_colombia = holidays.Colombia(years=anios)
        
        def contar_dias_habiles_colombia(fecha_inicio, fecha_fin):
            """Cuenta días laborales descartando sábados, domingos y festivos en Colombia."""
            if pd.isna(fecha_inicio) or pd.isna(fecha_fin) or fecha_inicio > fecha_fin:
                return 0
            dias = pd.date_range(start=fecha_inicio, end=fecha_fin, freq="D")
            dias_habiles = [
                d for d in dias 
                if d.weekday() < 5 and d.strftime("%Y-%m-%d") not in festivos_colombia
            ]
            return max(0, len(dias_habiles) - 1)
        
        def procesar_fallas_grupo(grupo_df):
            total_fallas = (grupo_df["Asistencia"] == "IS").sum()
            if total_fallas == 0:
                return pd.Series({"FACU": None})
            
            asistencias_con_fecha = grupo_df[grupo_df["Asistencia"].isin(["A", "IS", "AR"])][["Asistencia", "Fecha"]].values.tolist()
            
            max_consecutivas = 0
            contador_actual = 0
            fecha_2_consecutiva = None
            fecha_3_consecutiva = None
            
            for estado, fecha in asistencias_con_fecha:
                if estado == "IS":
                    contador_actual += 1
                    if contador_actual == 2 and fecha_2_consecutiva is None:
                        fecha_2_consecutiva = fecha
                    elif contador_actual == 3 and fecha_3_consecutiva is None:
                        fecha_3_consecutiva = fecha
                    
                    if contador_actual > max_consecutivas:
                        max_consecutivas = contador_actual
                else:
                    contador_actual = 0
            
            # Calcular días hábiles transcurridos desde cada hito de falla
            dh_2c = contar_dias_habiles_colombia(fecha_2_consecutiva.normalize(), fecha_actual) if fecha_2_consecutiva is not None else 0
            dh_3c = contar_dias_habiles_colombia(fecha_3_consecutiva.normalize(), fecha_actual) if fecha_3_consecutiva is not None else 0
            
            cumple_2c = dh_2c >= 5
            cumple_3c = dh_3c >= 5
            
            # Lógica de etiquetado progresivo
            if max_consecutivas >= 3:
                if cumple_3c:
                    facu_str = "3 C"
                elif cumple_2c:
                    # Racha 2C madura + 3.ª falla aún sin cumplir los 5 días hábiles
                    facu_str = f"2 C ({total_fallas})"
                else:
                    facu_str = f"{total_fallas}"
            elif max_consecutivas == 2:
                if cumple_2c:
                    facu_str = "2 C"
                else:
                    facu_str = f"{total_fallas}"
            else:
                facu_str = f"{total_fallas}"
                
            return pd.Series({"FACU": facu_str})

        df_fallas = (
            df_asistencia
            .groupby(cols_agrupacion, group_keys=False)
            .apply(procesar_fallas_grupo)
            .reset_index()
        )
        
        fallas_acumuladas = df_fallas[df_fallas["FACU"].notna()].reset_index(drop=True)
        return fallas_acumuladas

    except Exception as e:
        st.error(f"Error al calcular fallas acumuladas: {e}")
        return pd.DataFrame(columns=cols_base)


# ==========================================
# CARGA DE ARCHIVO MAESTRO DESDE GITHUB / LOCAL
# ==========================================
@st.cache_data(ttl=300)
def cargar_datos_excel():
    if os.path.exists(DB_FILE):
        return pd.read_excel(DB_FILE, sheet_name=None)
    
    if "GITHUB_TOKEN" in st.secrets:
        try:
            url = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{DB_FILE}"
            headers = {"Authorization": f"token {st.secrets['GITHUB_TOKEN']}"}
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                return pd.read_excel(io.BytesIO(response.content), sheet_name=None)
        except Exception as e:
            st.warning(f"No se pudo cargar desde GitHub: {e}")
    return None


# ==========================================
# INTERFAZ PRINCIPAL STREAMLIT
# ==========================================
st.title("📊 Sistema de Control de Asistencia y Evaluaciones - SENA")

dic_excel = cargar_datos_excel()

if dic_excel is None:
    st.error("⚠️ No se encontró la base de datos de Excel ('Reporte de Asistencia.xlsx'). Suba el archivo o verifique la conexión con GitHub.")
    st.stop()

# Selección de Pestaña principal en Excel (Grupos/Sedes)
hojas_disponibles = list(dic_excel.keys())
hoja_seleccionada = st.sidebar.selectbox("Seleccione la Hoja / Grupo Maestro", hojas_disponibles)
df_maestro = dic_excel[hoja_seleccionada]

# Filtros principales
instructores = sorted(df_maestro["Instructor"].dropna().astype(str).unique()) if "Instructor" in df_maestro.columns else []
instructor_sel = st.sidebar.selectbox("Instructor", instructores)

grupos = sorted(df_maestro[df_maestro["Instructor"] == instructor_sel]["Grupo"].dropna().astype(str).unique()) if instructor_sel else []
grupo_sel = st.sidebar.selectbox("Ficha / Grupo", grupos)

df_filtrado = df_maestro[(df_maestro["Instructor"] == instructor_sel) & (df_maestro["Grupo"].astype(str) == str(grupo_sel))]

# Calcular FACU actual
df_facu_actual = calcular_fallas_acumuladas()

# Pestañas de la Aplicación
tab1, tab2, tab3 = st.tabs(["📋 Llamado a Lista (Asistencia)", "📝 Registro de Evaluaciones", "📈 Histórico y Descargas"])

# ------------------------------------------
# TAB 1: LLAMADO A LISTA
# ------------------------------------------
with tab1:
    st.subheader(f"Control de Asistencia - Ficha: {grupo_sel}")
    
    # Alerta de Aprendices con 2C / 3C
    if not df_facu_actual.empty:
        alertas = df_facu_actual[
            (df_facu_actual["Grupo"].astype(str) == str(grupo_sel)) & 
            (df_facu_actual["FACU"].astype(str).str.contains("2 C|3 C", na=False))
        ]
        if not alertas.empty:
            st.warning("⚠️ **Alertas de Deserción / Fallas Consecutivas en este Grupo:**")
            st.dataframe(alertas[["Documento", "Nombre", "FACU"]], use_container_width=True)

    fecha_asistencia = st.date_input("Fecha de Asistencia", datetime.now())
    trimestre = st.text_input("Trimestre", "Trimestre I")
    asignacion = st.number_input("Número de Asignación / Sesión", min_value=1, value=1)
    resultados = st.text_input("Resultado de Aprendizaje / Tema", "General")

    if not df_filtrado.empty:
        df_asist_form = df_filtrado[["Documento", "Nombre"]].copy()
        
        # Unir con FACU actual para visualizar el estado previo al tomar lista
        if not df_facu_actual.empty:
            df_asist_form = df_asist_form.merge(
                df_facu_actual[df_facu_actual["Grupo"].astype(str) == str(grupo_sel)][["Documento", "FACU"]],
                on="Documento", how="left"
            )
        else:
            df_asist_form["FACU"] = None

        df_asist_form["Asistencia"] = "A"  # Valor por defecto: Asiste

        edited_df = st.data_editor(
            df_asist_form,
            column_config={
                "Asistencia": st.column_config.SelectboxColumn(
                    "Estado Asistencia",
                    options=["A", "IS", "AR"],
                    help="A: Asiste | IS: Inasistencia | AR: Arribo Tardío",
                    required=True
                )
            },
            disabled=["Documento", "Nombre", "FACU"],
            hide_index=True,
            use_container_width=True
        )

        if st.button("💾 Guardar Asistencia"):
            registros = []
            for _, row in edited_df.iterrows():
                registros.append({
                    "Fecha": fecha_asistencia.strftime("%Y-%m-%d"),
                    "Grupo": grupo_sel,
                    "Instructor": instructor_sel,
                    "Trimestre": trimestre,
                    "Asignacion_Num": asignacion,
                    "Resultados": resultados,
                    "Documento": row["Documento"],
                    "Nombre": row["Nombre"],
                    "Asistencia": row["Asistencia"]
                })
            
            df_nuevos = pd.DataFrame(registros)
            df_nuevos.to_csv("asistencia_guardada.csv", mode='a', header=not os.path.exists("asistencia_guardada.csv"), index=False)
            
            # Sincronización con GitHub
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
                    st.success("🔄 ¡Historial de Asistencias respaldado exitosamente en GitHub!")
                except Exception as e:
                    st.warning(f"Guardado localmente, pero ocurrió un error al subir a GitHub: {e}")
            else:
                st.success("✅ Asistencia guardada localmente.")
            
            st.rerun()

# ------------------------------------------
# TAB 2: REGISTRO DE EVALUACIONES
# ------------------------------------------
with tab2:
    st.subheader(f"Registro de Evaluaciones - Ficha: {grupo_sel}")
    
    fecha_eval = st.date_input("Fecha de Evaluación", datetime.now(), key="eval_fecha")
    trimestre_eval = st.text_input("Trimestre", "Trimestre I", key="eval_trim")
    materia_eval = st.text_input("Resultado / Competencia", "Competencia Técnica", key="eval_mat")

    if not df_filtrado.empty:
        df_eval_form = df_filtrado[["Documento", "Nombre"]].copy()
        df_eval_form["Evaluación (A/D)"] = "A"
        df_eval_form["Observaciones"] = ""

        edited_eval_df = st.data_editor(
            df_eval_form,
            column_config={
                "Evaluación (A/D)": st.column_config.SelectboxColumn(
                    "Juicio de Evaluación",
                    options=["A", "D"],
                    help="A: Aprobado | D: Deficiente / No Aprobado",
                    required=True
                )
            },
            disabled=["Documento", "Nombre"],
            hide_index=True,
            use_container_width=True
        )

        if st.button("💾 Guardar Evaluaciones"):
            registros_eval = []
            for _, row in edited_eval_df.iterrows():
                registros_eval.append({
                    "Fecha": fecha_eval.strftime("%Y-%m-%d"),
                    "Grupo": grupo_sel,
                    "Instructor": instructor_sel,
                    "Trimestre": trimestre_eval,
                    "Materia": materia_eval,
                    "Documento": row["Documento"],
                    "Nombre": row["Nombre"],
                    "Evaluación (A/D)": row["Evaluación (A/D)"],
                    "Observaciones": row["Observaciones"]
                })
            
            df_nuevas_eval = pd.DataFrame(registros_eval)
            df_nuevas_eval.to_csv("evaluaciones_guardadas.csv", mode='a', header=not os.path.exists("evaluaciones_guardadas.csv"), index=False)
            
            if "GITHUB_TOKEN" in st.secrets:
                try:
                    g = Github(st.secrets["GITHUB_TOKEN"])
                    repo = g.get_repo(REPO_NAME)
                    with open("evaluaciones_guardadas.csv", "r", encoding='utf-8') as f:
                        contenido_csv = f.read()
                    try:
                        sha = repo.get_contents("evaluaciones_guardadas.csv", ref="main").sha
                        repo.update_file("evaluaciones_guardadas.csv", "🤖 Actualizar histórico evaluaciones CSV", contenido_csv, sha, branch="main")
                    except Exception:
                        repo.create_file("evaluaciones_guardadas.csv", "🤖 Crear histórico evaluaciones CSV", contenido_csv, branch="main")
                    st.success("🔄 ¡Historial de Evaluaciones respaldado en GitHub!")
                except Exception as e:
                    st.warning(f"Guardado localmente, pero no se pudo subir a GitHub: {e}")
            else:
                st.success("✅ Evaluaciones guardadas localmente.")
            
            st.rerun()

# ------------------------------------------
# TAB 3: HISTÓRICO Y DESCARGAS
# ------------------------------------------
with tab3:
    st.subheader("📊 Consulta de Registros e Históricos")
    
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["Histórico Asistencia", "Histórico Evaluaciones", "Consolidado FACU"])
    
    with sub_tab1:
        if os.path.exists("asistencia_guardada.csv"):
            df_asist_hist = pd.read_csv("asistencia_guardada.csv")
            st.dataframe(df_asist_hist, use_container_width=True)
            
            csv_asist = df_asist_hist.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Descargar Historial de Asistencias (CSV)",
                data=csv_asist,
                file_name="asistencia_guardada_completa.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("No hay registros de asistencia guardados aún.")
            
    with sub_tab2:
        if os.path.exists("evaluaciones_guardadas.csv"):
            df_eval_hist = pd.read_csv("evaluaciones_guardadas.csv")
            st.dataframe(df_eval_hist, use_container_width=True)
            
            csv_eval = df_eval_hist.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Descargar Historial de Evaluaciones (CSV)",
                data=csv_eval,
                file_name="evaluaciones_guardadas_completa.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("No hay registros de evaluaciones guardados aún.")

    with sub_tab3:
        if not df_facu_actual.empty:
            st.write("#### Reporte Consolidado de Fallas Acumuladas (FACU)")
            st.dataframe(df_facu_actual, use_container_width=True)
            
            csv_facu = df_facu_actual.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Descargar Reporte FACU (CSV)",
                data=csv_facu,
                file_name="reporte_facu_consolidado.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("No hay fallas registradas que requieran alerta FACU.")
