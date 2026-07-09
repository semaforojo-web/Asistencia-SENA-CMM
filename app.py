import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuración de la página para entorno móvil y de escritorio
st.set_page_config(page_title="Control de Asistencia y Evaluación - SENA", layout="wide")

# Nombre del archivo original de Excel en tu repositorio de GitHub
DB_FILE = "Reporte de Asistencia.xlsx"

def cargar_datos():
    """Carga y procesa el listado de aprendices desde la hoja correspondiente"""
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_excel(DB_FILE, sheet_name="Listado de aprendices", header=None)
            if df.empty:
                return pd.DataFrame(columns=["Ficha", "Documento", "Nombre Completo"])
                
            df_procesado = pd.DataFrame()
            
            # Columna E (Ficha) -> Posición 4
            df_procesado["Ficha"] = df.iloc[:, 4].astype(str).str.strip()
            # Columna L (Documento) -> Posición 11
            df_procesado["Documento"] = df.iloc[:, 11].fillna("S/D").astype(str).str.strip()
            # Columnas M a O (Nombres y Apellidos) -> Posiciones 12 a 14
            nombres_completos = df.iloc[:, 12:15].fillna("").astype(str)
            df_procesado["Nombre Completo"] = nombres_completos.agg(' '.join, axis=1).str.replace(r'\s+', ' ', regex=True).str.strip()
            
            # FILTRADO POR COLUMNA P (Posición 15) - Estados de deserción o retiro
            df_procesado["Estado"] = df.iloc[:, 15].fillna("").astype(str).str.upper().str.strip()
            terminos_excluir = "CANCELADO|RETIRO VOLUNTARIO|TRASLADO"
            df_procesado = df_procesado[~df_procesado["Estado"].str.contains(terminos_excluir, regex=True)]
            
            # Asegurar que la ficha sea numérica válida
            df_procesado = df_procesado[df_procesado["Ficha"].str.contains(r'^\d+$', na=False)]
            df_procesado["Nombre Completo"] = df_procesado["Nombre Completo"].replace("", "Aprendiz sin nombre registrado")
            
            return df_procesado[["Ficha", "Documento", "Nombre Completo"]].reset_index(drop=True)
        except Exception as e:
            st.error(f"Error al filtrar listado de aprendices: {e}")
            
    data = {"Ficha": ["Error"], "Documento": ["0"], "Nombre Completo": ["ARCHIVO EXCEL NO DETECTADO"]}
    return pd.DataFrame(data)

def obtener_instructores_desde_lista():
    """Busca los nombres de los instructores en la hoja dedicada 'Listado de instructores' (Columna A)"""
    if os.path.exists(DB_FILE):
        try:
            df_inst = pd.read_excel(DB_FILE, sheet_name="Listado de instructores", header=None)
            # Tomamos la primera columna (Posición 0)
            instructores = df_inst.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist()
            # Limpieza de posibles encabezados repetidos
            instructores = [i for i in instructores if i.upper() != "INSTRUCTOR" and i != "" and i.upper() != "NAN"]
            return sorted(instructores)
        except Exception:
            return ["Falta hoja 'Listado de instructores'"]
    return ["No Detectado"]

def obtener_trimestres_disponibles(ficha, instructor):
    """Busca los trimestres vinculados a una ficha e instructor en el Cabezote"""
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            
            filtro = (df_cab[6] == str(ficha)) & (df_cab[5] == str(instructor).upper())
            resultado = df_cab[filtro]
            
            if not resultado.empty and resultado.shape[1] > 47:
                trimestres = resultado.iloc[:, 47].dropna().astype(str).str.strip().unique().tolist()
                trimestres_validos = sorted([t for t in trimestres if t != "" and t.upper() != "NAN" and t.upper() != "TRIMESTRE"])
                if trimestres_validos:
                    return trimestres_validos
        except Exception:
            pass
    return ["Sin trimestres detectados"]

def obtener_materias_disponibles(ficha, instructor, trimestre):
    """Obtiene los números de materia cargados para la combinación seleccionada"""
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            df_cab[47] = df_cab[47].astype(str).str.strip() if df_cab.shape[1] > 47 else ""
            
            filtro = (df_cab[6] == str(ficha)) & \
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

def filtrar_materia_final(ficha, instructor, trimestre, materia_num):
    """Retorna el nombre largo de la asignatura basándose en el número seleccionado"""
    if os.path.exists(DB_FILE):
        try:
            df_cab = pd.read_excel(DB_FILE, sheet_name="Cabezote", header=None)
            df_cab[3] = df_cab[3].astype(str).str.strip()
            df_cab[5] = df_cab[5].astype(str).str.strip().str.upper()
            df_cab[6] = df_cab[6].astype(str).str.strip()
            df_cab[47] = df_cab[47].astype(str).str.strip() if df_cab.shape[1] > 47 else ""
            
            filtro = (df_cab[6] == str(ficha)) & \
                     (df_cab[5] == str(instructor).upper()) & \
                     (df_cab[47] == str(trimestre)) & \
                     (df_cab[3] == str(materia_num))
                     
            resultado = df_cab[filtro]
            
            if not resultado.empty:
                materia_texto = resultado.iloc[0, 10] if resultado.shape[1] > 10 else resultado.iloc[0, 3]
                if pd.notna(materia_texto) and str(materia_texto).strip() != "":
                    return str(materia_texto).strip()
            return f"Materia {materia_num} (Sin descripción en Cabezote)"
        except Exception as e:
            return f"Error: {e}"
    return "Archivo Excel no encontrado"

# --- CONFIGURACIÓN DE DATOS INICIALES ---
df_aprendices = cargar_datos()
lista_instructores = obtener_instructores_desde_lista()

# Creación de históricos locales si no existen
if not os.path.exists("asistencia_guardada.csv"):
    pd.DataFrame(columns=["Fecha", "Ficha", "Instructor", "Trimestre", "Materia_Num", "Materia_Nombre", "Documento", "Nombre", "Asistencia"]).
            
