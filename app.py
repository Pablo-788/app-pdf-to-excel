import streamlit as st
from auth import procesar_callback, cerrar_sesion
from ui import mostrar_login, render_header, render_footer, mostrar_aplicacion
from cron import iniciar_cron

def main():
    if procesar_callback():
        return
    if "access_token" in st.session_state:
        mostrar_aplicacion()
    else:
        mostrar_login()
    if st.query_params.get("logout"):
        cerrar_sesion()
        return

if __name__ == "__main__":
    iniciar_cron()
    main()