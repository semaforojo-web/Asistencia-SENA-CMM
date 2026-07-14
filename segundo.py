import streamlit as st
from github import Github

def verificar_conexion():
    try:
        # 1. Autenticación
        g = Github(st.secrets["GITHUB_TOKEN"])
        user = g.get_user()
        st.write(f"✅ Autenticado como: {user.login}")

        # 2. Intentar acceder al repositorio
        repo = g.get_repo("semaforojo-web/Asistencia-SENA-CMM")
        
        # 3. Verificar permisos de escritura de forma directa
        # Si el usuario tiene acceso de escritura, esto no debería fallar.
        # Si el token no tiene permisos suficientes, lanzará una excepción aquí.
        st.write(f"✅ Repositorio encontrado: {repo.full_name}")
        st.write(f"Permisos del usuario según GitHub: {repo.permissions}")

        if repo.permissions.push:
            st.success("🎉 ¡El token tiene permisos de ESCRITURA (Push)! El error no es de permisos.")
        else:
            st.error("❌ El token NO tiene permisos de escritura (Push). Debes revisar tu Personal Access Token.")
            
    except Exception as e:
        st.error(f"❌ Error al conectar o validar: {e}")

verificar_conexion()
