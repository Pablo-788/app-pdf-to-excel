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
    """
    Sube un archivo a SharePoint usando la API de Microsoft Graph directamente.
    """
    access_token = st.session_state.get("access_token")
    if not access_token:
        st.error("Error de autenticación: No se encontró el token de acceso.")
        return False

    # --- Configuración de tu SharePoint ---
    HOST_NAME = "saboraespana.sharepoint.com"
    SITE_PATH = "/sites/departamento.ti"
    
    # --- CORRECCIÓN DEFINITIVA DE LA RUTA ---
    # La ruta debe ser RELATIVA a la biblioteca de documentos principal ("Documentos compartidos").
    # Por lo tanto, no incluimos "Documentos compartidos" en la ruta.
    FOLDER_PATH = "General/PoC Plantillas SaEGA"

    # --- Codificación de caracteres especiales en el nombre del archivo ---
    nombre_archivo_encoded = urllib.parse.quote(nombre_archivo)

    # --- Construcción de la URL para la API de Graph (VERSIÓN FINAL) ---
    site_identifier = f"{HOST_NAME}:{SITE_PATH}"
    
    # La URL ahora usa la ruta corregida
    graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_identifier}/drive/root:/{FOLDER_PATH}/{nombre_archivo_encoded}:/content"

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }

    try:
        bytes_io.seek(0)
        file_content = bytes_io.getvalue()
        
        response = requests.put(url=graph_url, headers=headers, data=file_content)
        
        response.raise_for_status()
        
        return True
        
    except requests.exceptions.HTTPError as e:
        error_details = e.response.json()
        error_message = error_details.get("error", {}).get("message", "Sin detalles adicionales.")
        st.error(f"❌ Error al subir a SharePoint: {e.response.status_code} - {error_message}")
        return False
    except Exception as e:
        st.error(f"❌ Ocurrió un error inesperado: {e}")
        return False