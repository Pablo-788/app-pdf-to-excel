# auth.py
import os
from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components
import msal

import jwt

load_dotenv()

CLIENT_ID  = os.getenv("CLIENT_ID")
TENANT_ID  = os.getenv("TENANT_ID")
AUTHORITY    = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:8501/"

# Opción A: usar /.default (requiere admin consent previo de los permisos delegados configurados)
SCOPES = ["https://graph.microsoft.com/.default"]


# Opción B (recomendada para pruebas): pedir explícitamente los scopes delegados que necesitas
# SCOPES = ["Files.ReadWrite.All", "Sites.ReadWrite.All", "offline_access", "openid", "profile", "email"]

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

def _decode_access_token(token: str) -> dict:
    try:
        return jwt.decode(token, options={"verify_signature": False, "verify_aud": False})
    except Exception as e:
        return {"_decode_error": str(e)}

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
        st.session_state.id_token     = result.get("id_token")
        st.session_state.user_info    = result.get("id_token_claims", {})

  
        st.session_state.token_claims = _decode_access_token(result["access_token"])
  
        claims = st.session_state.token_claims
        st.session_state.token_mode = "delegated" if claims.get("scp") else ("app-only" if claims.get("roles") else "unknown")

 
        scp   = claims.get("scp")
        roles = claims.get("roles")
        resumen = f"modo={st.session_state.token_mode} | scp={scp} | roles={roles}"
        st.toast(f"Token obtenido: {resumen}")

        components.html("<script>window.close();</script>", height=0, width=0)
        st.query_params.clear()
        st.rerun()
        return True

    st.error(f"❌ No se pudo obtener el token:\n{result.get('error_description')}")
    return False

def cerrar_sesion():
    st.query_params.clear()
    for k in ("access_token", "id_token", "user_info", "token_claims", "token_mode", "msal_app"):
        st.session_state.pop(k, None)
    st.rerun()
