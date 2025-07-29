import streamlit as st
from extraer_tabla import procesar_pdf
import threading   
import subprocess
import time
import os
import requests
from urllib.parse import urlencode

st.set_page_config(page_title="Convertidor PDF → Excel", layout="centered")

# Variables de entorno
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")
REDIRECT_URI = os.environ.get("REDIRECT_URI")
SCOPE = "openid offline_access https://graph.microsoft.com/User.Read"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0"
SCOPE = "User.Read"

def get_auth_url():
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": SCOPE,
        "state": "12345"  # Para simplificar, luego mejorarlo
    }
    url = f"{AUTHORITY}/authorize?{urlencode(params)}"
    return url

# --- MANEJO DE LOGIN OAUTH ---

query_params = st.query_params

# 1. Si ya tenemos token guardado, no hacemos nada
if "access_token" not in st.session_state:

    # 2. Si viene un ?code en la URL (Microsoft nos redirigió)
    if "code" in query_params:
        code = query_params["code"][0]

        token_url = f"{AUTHORITY}/token"
        data = {
            "client_id": CLIENT_ID,
            "scope": SCOPE,
            "code": code,
            "redirect_uri": REDIRECT_URI,
            "grant_type": "authorization_code",
            "client_secret": CLIENT_SECRET
        }

        response = requests.post(token_url, data=data)
        tokens = response.json()

        if "access_token" not in tokens:
            st.error("Error al obtener el token de acceso:")
            st.write(tokens)
            st.stop()

        st.session_state["access_token"] = tokens["access_token"]
        st.session_state["logged_in"] = True

        # Limpia la URL de ?code y recarga la app
        st.query_params.clear()
        st.experimental_rerun()

    else:
        # 3. No hay sesión ni código => mostrar botón de login
        st.title("Iniciar sesión")
        auth_url = get_auth_url()
        st.markdown(f"[Inicia sesión con Microsoft 365]({auth_url})")
        st.stop()

# Botón para cerrar sesión
with st.sidebar:
    st.markdown("### Opciones")
    if st.button("🔒 Cerrar sesión"):
        st.session_state.clear()
        st.experimental_rerun()


st.title("Convertidor de PDF a Excel")
st.markdown("Sube un archivo PDF de factura para convertirlo automáticamente a Excel.")

# Subida del archivo PDF
pdf_file = st.file_uploader("Selecciona un archivo PDF", type=["pdf"])

if pdf_file:
    with st.spinner("Procesando..."):
        try:
            output_excel, nombre_archivo = procesar_pdf(pdf_file, pdf_file.name)
            st.success("¡Conversión completada!")

            st.download_button(
                label="📥 Descargar Excel",
                data=output_excel,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el PDF: {e}")

def cron_loop():
    while True:
        print("Ejecutando petición...")
        subprocess.call(["/bin/bash", "ping.sh"])
        time.sleep(600)  # 10 minutos
 
# Lanzar hilo del cron
threading.Thread(target=cron_loop, daemon=True).start()