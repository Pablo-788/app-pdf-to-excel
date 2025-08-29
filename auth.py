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
SCOPES       =  [
    "User.Read",          # Perfil básico del usuario
    "Files.Read",         # Leer archivos del usuario en OneDrive y SharePoint
    "Sites.Read.All",     # Leer cualquier sitio que el usuario pueda ver
    "Files.Read.All",     # Leer todos los archivos de los usuarios a los que el usuario tiene acceso
    "Files.ReadWrite",    # Leer y escribir archivos del usuario en OneDrive y SharePoint
    "Sites.ReadWrite.All" # Leer y escribir en todos los sitios que el usuario puede ver
]
ALLOWED_GROUP_ID = os.getenv("ALLOWED_GROUP_ID")

def get_msal_app():
    """Devuelve la instancia de MSAL y la crea si no existe."""
    if "msal_app" not in st.session_state:
        st.session_state.msal_app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY
        )
    return st.session_state.msal_app

# ---------- Flujo MSAL ----------
def iniciar_autenticacion() -> str:
    msal_app = get_msal_app()
    """Devuelve la URL a la página de login de Microsoft."""
    return st.session_state.msal_app.get_authorization_request_url(
        SCOPES,
        redirect_uri=REDIRECT_URI
    )

def procesar_callback() -> bool:
    msal_app = get_msal_app()
    if "access_token" in st.session_state or "code" not in st.query_params:
        return False

    code = st.query_params["code"]
    if isinstance(code, list):
        code = code[0]

    result = st.session_state.msal_app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )

    if "access_token" in result:
        user_claims = result.get("id_token_claims", {})
        user_groups = user_claims.get("groups") or []

        if ALLOWED_GROUP_ID not in user_groups:
            st.error("❌ No tienes permisos para acceder a esta aplicación.")
            return True  # Detiene el flujo y evita mostrar la app

        # Si está en el grupo permitido, guarda sesión
        st.session_state.access_token = result["access_token"]
        st.session_state.user_info    = user_claims

        # Limpiar parámetros y recargar la app principal
        st.query_params.clear()
        st.rerun()
        return True

    st.error(f"❌ No se pudo obtener el token:\n{result.get('error_description')}")
    return False

def cerrar_sesion():
    """Limpia la sesión y recarga la aplicación."""
    st.query_params.clear()
    # Elimina todos los datos de sesión, incluyendo el objeto msal_app
    for k in ("access_token", "user_info", "msal_app"):
        st.session_state.pop(k, None)
    st.rerun()