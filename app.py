import streamlit as st
from auth import procesar_callback, cerrar_sesion
from ui import mostrar_login, render_header, render_footer, mostrar_aplicacion
from cron import iniciar_cron

def main():
    # El callback de autenticación se procesa primero
    if procesar_callback():
        return

    # Muestra la aplicación o la página de login según el estado de la sesión
    if "access_token" in st.session_state:
        mostrar_aplicacion()
    else:
        mostrar_login()

if __name__ == "__main__":
    iniciar_cron()
    main()
