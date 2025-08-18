import os
from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components
import msal

CLIENT_ID  = os.getenv("CLIENT_ID",  "your-client-id")   # mismo efecto que .get()
TENANT_ID  = os.getenv("TENANT_ID",  "your-tenant-id")
AUTHORITY    = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "https://app-pdf-to-excel.onrender.com/"            # URI registrada en Mobile & desktop
SCOPES       = ["User.Read"]
ALLOWED_GROUP_ID = os.getenv("ALLOWED_GROUP_ID")

def get_msal_app():
    """Devuelve la instancia de MSAL y la crea si no existe."""
    if "msal_app" not in st.session_state:
        st.session_state.msal_app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY
        )
    return st.session_state.msal_app

# ---------- Helper ----------
def abrir_en_nueva_pestana(url: str):
    """Abre la URL en un tab secundario."""
    components.html(
        f"""
        <script>
            window.open("{url}", "_blank");
        </script>
        """,
        height=0,
        width=0
    )

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
    """
    Si la URL trae ?code=..., lo intercambia por un access_token.
    Cierra la pestaña hija y recarga la principal.
    """
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

        # --- cerrar la pestaña (si es hija) ---
        components.html(
            """
            <script>
                window.close();
            </script>
            """,
            height=0,
            width=0
        )
        # Limpiar parámetros y recargar la app principal
        st.query_params.clear()
        st.rerun()
        return True

    st.error(f"❌ No se pudo obtener el token:\n{result.get('error_description')}")
    return False

def cerrar_sesion():
    st.st.experimental_set_query_params()
    for k in ("access_token", "user_info"):
        st.session_state.pop(k, None)
    st.rerun()