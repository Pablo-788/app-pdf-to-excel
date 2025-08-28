# sharepoint_utils.py
import streamlit as st
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
# IMPORTACIÓN CLAVE Y CORRECTA
from office365.runtime.auth.authentication_context import AuthenticationContext
import requests
import urllib.parse

def get_sharepoint_context() -> ClientContext:
    """
    Crea y devuelve un ClientContext de SharePoint autenticado.
    Esta es la implementación canónica y correcta para usar un token existente.
    """
    sharepoint_url = "https://saboraespana.sharepoint.com/sites/departamento.ti"
    
    # --- LA SOLUCIÓN CORRECTA Y DEFINITIVA ---
    
    # Paso 1: Crear un contexto de autenticación vacío.
    auth_ctx = AuthenticationContext(url=sharepoint_url)
    
    # Paso 2: Definir una función que devuelva el token.
    # La librería llamará a esta función cuando necesite autenticarse.
    def acquire_token_callback():
        return {'access_token': st.session_state.get("access_token")}

    # Paso 3: "Enganchar" la función de adquisición de token al contexto de autenticación.
    auth_ctx.acquire_token_func = acquire_token_callback
    
    # Paso 4: Crear el ClientContext final pasándole el contexto de autenticación ya listo.
    ctx = ClientContext(sharepoint_url, auth_ctx)
    
    # --- FIN DE LA SOLUCIÓN ---
    
    return ctx

def subir_a_sharepoint(bytes_io: BytesIO, nombre_archivo: str) -> bool:
    import urllib.parse, requests, streamlit as st

    access_token = st.session_state.get("access_token")
    if not access_token:
        st.error("Error de autenticación: No se encontró el token de acceso.")
        return False

    HOST_NAME  = "saboraespana.sharepoint.com"
    SITE_PATH  = "/sites/departamento.ti"              # ¡con la barra inicial!
    FOLDER_PATH = "General/testPlantillas"       # sin “Documentos compartidos”

    # Codifica carpeta y nombre (mantén las / de la ruta)
    encoded_folder = urllib.parse.quote(FOLDER_PATH.strip("/"), safe="/")
    encoded_name   = urllib.parse.quote(nombre_archivo)

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    try:
        # 1) Resolver el siteId por ruta
        site_url = f"https://graph.microsoft.com/v1.0/sites/{HOST_NAME}:{SITE_PATH}"
        site_resp = requests.get(site_url, headers=headers)
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]

        # 2) Subir usando el siteId (drive por defecto del sitio)
        upload_url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/drive/root:/{encoded_folder}/{encoded_name}:/content"
        )

        bytes_io.seek(0)
        put_resp = requests.put(upload_url, headers=headers, data=bytes_io.getvalue())
        put_resp.raise_for_status()
        return True

    except requests.exceptions.HTTPError as e:
        try:
            err = e.response.json()
            msg = err.get("error", {}).get("message", str(err))
        except Exception:
            msg = e.response.text
        st.error(f"❌ Error al subir a SharePoint: {e.response.status_code} - {msg}")
        return False
    except Exception as e:
        st.error(f"❌ Ocurrió un error inesperado: {e}")
        return False
