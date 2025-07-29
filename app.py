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

# --- LOGIN ---

query_params = st.query_params

if "code" not in query_params:
    st.title("Iniciar sesión")
    auth_url = get_auth_url()
    st.markdown(f"[Inicia sesión con Microsoft 365]({auth_url})")
    st.stop()  # No mostrar nada más hasta que se inicie sesión

# --- CALLBACK ---

else:
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
        st.error("Error obteniendo token:")
        st.write(tokens)
        st.stop()

    access_token = tokens["access_token"]

    st.success("¡Login correcto!")
    st.write(f"Token de acceso: {access_token[:40]}...")


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
        subprocess.call(["/bin/bash", "/app/ping.sh"])
        time.sleep(600)  # 10 minutos
 
# Lanzar hilo del cron
threading.Thread(target=cron_loop, daemon=True).start()