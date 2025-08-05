# app.py – Convertidor PDF ➔ Excel con login Microsoft (pestaña hija se cierra sola)
from concurrent.futures import process
import os
from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components
import msal
from extraer_tabla import procesar_pdf            # tu extractor de tablas
import threading, subprocess, time

# ---------- Configuración ----------
st.set_page_config(page_title="Convertidor PDF → Excel", layout="centered")
load_dotenv()

CLIENT_ID  = os.getenv("CLIENT_ID",  "your-client-id")   # mismo efecto que .get()
TENANT_ID  = os.getenv("TENANT_ID",  "your-tenant-id")
AUTHORITY    = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "https://app-pdf-to-excel.onrender.com/"            # URI registrada en Mobile & desktop
SCOPES       = ["User.Read"]
ALLOWED_GROUP_ID = os.getenv("ALLOWED_GROUP_ID")

# Crear instancia MSAL una sola vez
if "msal_app" not in st.session_state:
    st.session_state.msal_app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY
    )

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
    """Devuelve la URL a la página de login de Microsoft."""
    return st.session_state.msal_app.get_authorization_request_url(
        SCOPES,
        redirect_uri=REDIRECT_URI
    )

def procesar_callback() -> bool:
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
    for k in ("access_token", "user_info"):
        st.session_state.pop(k, None)
    st.rerun()

# ---------- Interfaz ----------
def mostrar_login():
    st.title("🔐 Iniciar Sesión")
    st.markdown("### Convertidor de PDF a Excel\nInicia sesión con tu cuenta de Microsoft para continuar.")

    if st.button("🚀 Iniciar sesión con Microsoft", type="primary"):
        auth_url = iniciar_autenticacion()
        abrir_en_nueva_pestana(auth_url)
        st.info("Se abrió una pestaña nueva para autenticarte. Vuelve aquí cuando termines.")

def mostrar_aplicacion():
    # -- Sidebar --
    with st.sidebar:
        info = st.session_state.get("user_info", {})
        st.markdown("### 👤 Usuario")
        st.write(f"**Nombre:** {info.get('name', 'N/A')}")
        st.write(f"**Email:**  {info.get('preferred_username', 'N/A')}")
        st.markdown("---")
        if st.button("🔒 Cerrar sesión", use_container_width=True):
            cerrar_sesion()

    # -- Cuerpo --
    st.title("📄 ➡️ 📊 Convertidor de Facturas PDF a Excel")
    st.markdown("Sube un archivo PDF de factura y te devolveré un Excel con los datos.")

    with st.expander("ℹ️ Formato de factura esperado"):
        st.markdown("""
        **Extraigo automáticamente:**

        - Códigos de artículo y cantidades  
        - Información de tienda `TIENDA XXX`  
        - Datos y resumen del pedido
        """)

    pdf_file = st.file_uploader(
        "Selecciona un PDF de factura",
        type=["pdf"],
        help="Sube un PDF que contenga tablas con productos y la etiqueta 'TIENDA'."
    )

    if pdf_file and st.button("🔄 Procesar Factura", type="primary"):
        with st.spinner("Extrayendo datos, dame unos segundos…"):
            try:
                output_excel, nombre_archivo = procesar_pdf(pdf_file, pdf_file.name)
                st.success("✅ ¡Factura procesada!")
                st.download_button(
                    "📥 Descargar Excel",
                    data=output_excel.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            except Exception as e:
                st.error("❌ Error al procesar la factura.")
                with st.expander("Detalles del error"):
                    st.code(str(e))

# ---------- Cron opcional ----------
def iniciar_cron():
    def cron_loop():
        while True:
            try:
                subprocess.call(["/bin/bash", "ping.sh"], timeout=30)
            except Exception as e:
                print("Error en cron:", e)
            time.sleep(600)
    if "cron_started" not in st.session_state:
        threading.Thread(target=cron_loop, daemon=True).start()
        st.session_state.cron_started = True

# ---------- Main ----------
def main():
    if procesar_callback():              # Procesa ?code=...
        return
    if "access_token" in st.session_state:
        mostrar_aplicacion()
    else:
        mostrar_login()

if __name__ == "__main__":
    iniciar_cron()
    main()
