# auth.py
import os
from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components
import msal

load_dotenv()

CLIENT_ID  = os.getenv("CLIENT_ID")
TENANT_ID  = os.getenv("TENANT_ID")
AUTHORITY    = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:8501/"

# --- CAMBIO IMPORTANTE Y DEFINITIVO ---
# Este es el permiso correcto para la API de Microsoft Graph
SCOPES = ["https://graph.microsoft.com/.default"]

def get_msal_app():
    if "msal_app" not in st.session_state:
        st.session_state.msal_app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY
        )
    return st.session_state.msal_app

def iniciar_autenticacion() -> str:
    return get_msal_app().get_authorization_request_url(
        SCOPES,
        redirect_uri=REDIRECT_URI
    )

def procesar_callback() -> bool:
    if "access_token" in st.session_state or "code" not in st.query_params:
        return False

    code = st.query_params["code"]
    if isinstance(code, list):
        code = code[0]

    result = get_msal_app().acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )

    if "access_token" in result:
        st.session_state.access_token = result["access_token"]
        st.session_state.user_info    = result.get("id_token_claims", {})
        
        components.html("<script>window.close();</script>", height=0, width=0)
        st.query_params.clear()
        st.rerun()
        return True

    st.error(f"❌ No se pudo obtener el token:\n{result.get('error_description')}")
    return False

def cerrar_sesion():
    st.query_params.clear()
    for k in ("access_token", "user_info", "msal_app"):
        st.session_state.pop(k, None)
    st.rerun()