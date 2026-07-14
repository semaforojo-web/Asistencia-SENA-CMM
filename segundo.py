import streamlit as st
from github import Github

def verificar_conexion():
    try:
        # 1. Autenticación
        g = Github(st.secrets["GITHUB_TOKEN"])
        user = g.get_user()
        
        # 2. Verificación de Repositorio
        repo = g.get_repo("semaforojo-web/Asistencia-SENA-CMM")
        
        # 3. Verificación de permisos del usuario actual
        # permission_dict devolverá valores como {'admin': True, 'push': True, 'pull': True}
        permisos = repo.get_collaborator(user.login).permissions
        
        st.success(f"✅ Autenticado como: {user.login}")
        st.write("Permisos detectados:", permisos)
        
        if permisos.push:
            st.info("El token tiene permisos de ESCRITURA (Push).")
        else:
            st.error("❌ El token NO tiene permisos de escritura. Revisa el Personal Access Token.")
            
    except Exception as e:
        st.error(f"❌ Error al conectar o validar permisos: {e}")

verificar_conexion()
