import os
from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components
import msal
from extraer_tabla import procesar_pdf            # tu extractor de tablas
import threading, subprocess, time
import logging
from datetime import datetime
import pandas as pd
from io import BytesIO
import base64

# ---------- Configuración ----------
st.set_page_config(page_title="Convertidor Pedidos ET → Excel", layout="centered")
load_dotenv()

APP_TITLE   = "Convertidor Pedidos ET → Excel"    # cambia el texto cuando quieras
APP_VERSION = "0.3.16"

PRIMARY_COLOR     = "#0d6efd"   # azul principal (botones)
BRAND_BLUE_LIGHT  = "#E6F0FA"   # fondo barra superior
BRAND_BLUE_DARK   = "#003366"   # texto barra superior

# Coloca tu logo en el repositorio y apunta aquí:
LOGO_PATH = "LOGO_SAE.png"

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
    st.query_params.clear()
    for k in ("access_token", "user_info"):
        st.session_state.pop(k, None)
    st.rerun()

def inject_styles():
    st.markdown(f"""
    <style>
    /* ======= NAVBAR ======= */
    .navbar {{ display:flex; justify-content:space-between; align-items:center; 
               background:{BRAND_BLUE_LIGHT}; padding:10px 24px; }}
    .navbar-left {{ display:flex; align-items:center; gap:12px; }}
    .navbar-left img {{ height:40px; }}
    .navbar-title {{ font-weight:600; color:{BRAND_BLUE_DARK}; font-size:18px; }}
    .navbar-right {{ display:flex; align-items:center; gap:12px; }}

    .menu {{ position:relative; }}
    .menu-btn {{ background:transparent; border:none; cursor:pointer; 
                 font-size:16px; color:{BRAND_BLUE_DARK}; }}
    .menu-panel {{ display:none; position:absolute; right:0; top:36px; background:#fff;
                   min-width:220px; padding:12px; border-radius:10px; 
                   box-shadow:0 8px 24px rgba(0,0,0,.15); z-index:9999; }}
    .menu-panel p {{ margin:6px 0; }}
    .logout-btn {{ background:{PRIMARY_COLOR}; color:#fff; border:none; 
                   padding:8px 12px; border-radius:8px; cursor:pointer; width:100%; }}

    /* ======= CENTRADO ======= */
    .center-wrap {{
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        gap:16px;
        margin-top:40px;  /* opcional, se puede ajustar según contexto */
    }}

    .center-preview {{
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        gap:12px;
        margin-top:20px;
        width:100%;
    }}

    .recuadro-login {{
        background-color: #bc9a5f;  /* fondo negro */
        border-radius: 16px;     /* bordes redondeados */
        padding: 40px;            /* espacio interno */
        max-width: 500px;         /* ancho máximo */
        margin: 50px auto;        /* centrado horizontal y algo de margen arriba */
        text-align: center;       /* centrado de textos y botones */
        color: #fff;              /* texto blanco */
    }}

    /* ======= BOTONES ======= */
    .stButton > button,
    .stDownloadButton > button {{
        background:{PRIMARY_COLOR} !important;
        color:#fff !important; 
        border: 0 !important;
        border-radius: 8px !important;
        padding: 10px 16px !important;
        font-size: 16px !important;
    }}

    /* ======= FOOTER ======= */
    .footer {{ text-align:center; margin:48px 0 16px; color:#111; }}
    .footer .brand {{ color:{PRIMARY_COLOR}; font-weight:700; }}
    
    </style>
    """, unsafe_allow_html=True)


def render_header():
    inject_styles()

    # Agregamos estilos para animaciones suaves
    st.markdown(f"""
    <style>
    /* Animación para desplegables */
    .menu-panel {{
        overflow: hidden;
        max-height: 0;
        opacity: 0;
        transition: max-height 0.3s ease, opacity 0.3s ease;
    }}
    .menu-panel.open {{
        max-height: 500px; /* suficiente para que quepa el contenido */
        opacity: 1;
    }}
    </style>
    """, unsafe_allow_html=True)

    # Cargar logo en base64 (si no existe, no muestra imagen)
    logo_b64 = cargar_logo_base64()
    info = st.session_state.get("user_info", {})
    name  = info.get("name", "Usuario")
    email = info.get("preferred_username", "usuario@example.com")

    st.markdown(f"""
    <div class="navbar">
        <div class="navbar-left">
            {f'<img src="data:image/png;base64,{logo_b64}" alt="logo" />' if logo_b64 else ''}
            <div class="navbar-title">{APP_TITLE}</div>
        </div>
        <div class="navbar-right">
            <div class="menu">
                <button class="menu-btn" id="userBtn">{name}</button>
                <div class="menu-panel" id="userPanel">
                    <p><b>Nombre:</b> {name}</p>
                    <p><b>Email:</b> {email}</p>
                </div>
            </div>
            <div class="menu">
                <button class="menu-btn" id="profileBtn">👤</button>
                <div class="menu-panel" id="profilePanel">
                    <button class="logout-btn" id="logoutBtn">Cerrar sesión</button>
                    <p style="margin-top:8px; font-size:12px; opacity:.7;">Versión: {APP_VERSION}</p>
                </div>
            </div>
        </div>
    </div>

    <script>
    document.addEventListener("DOMContentLoaded", function() {{
        const userBtn = document.getElementById('userBtn');
        const userPanel = document.getElementById('userPanel');
        const profileBtn = document.getElementById('profileBtn');
        const profilePanel = document.getElementById('profilePanel');

        function toggle(panel) {{
            if (!panel) return;
            panel.classList.toggle('open');
        }}

        if (userBtn) userBtn.onclick = (e) => {{
            e.stopPropagation();
            toggle(userPanel);
            if (profilePanel) profilePanel.classList.remove('open');
        }};
        if (profileBtn) profileBtn.onclick = (e) => {{
            e.stopPropagation();
            toggle(profilePanel);
            if (userPanel) userPanel.classList.remove('open');
        }};

        document.addEventListener('click', function() {{
            if (userPanel) userPanel.classList.remove('open');
            if (profilePanel) profilePanel.classList.remove('open');
        }});

        const logoutBtn = document.getElementById('logoutBtn');
        if (logoutBtn) {{
            logoutBtn.onclick = function() {{
                const url = new URL(window.location.href);
                url.searchParams.set('logout', '1');
                window.location.href = url.toString();
            }};
        }}
    }});
    </script>
    """, unsafe_allow_html=True)


def render_footer():
    st.markdown(f"""
    <div class="footer">
        Made with sweetness by <span class="brand">SaE Tech Team!</span><br/>
        Version: {APP_VERSION}
    </div>
    """, unsafe_allow_html=True)

def render_login_navbar():
    # Cargar logo en base64 (si existe)
    logo_b64 = cargar_logo_base64()

    st.markdown(f"""
    <div style="
        display:flex; 
        align-items:center; 
        gap:12px; 
        background:{BRAND_BLUE_LIGHT}; 
        padding:10px 24px;
        border-bottom:1px solid #ccc;
    ">
        {f'<img src="data:image/png;base64,{logo_b64}" alt="Sabor a España" style="height:40px;" />' if logo_b64 else ''}
        <div style="font-weight:600; color:{BRAND_BLUE_DARK}; font-size:18px;">
            {APP_TITLE}
        </div>
    </div>
    """, unsafe_allow_html=True)

def cargar_logo_base64(ruta_logo: str = LOGO_PATH) -> str:
    """Carga el logo desde archivo y devuelve su contenido en base64. Si falla, devuelve cadena vacía."""
    try:
        with open(ruta_logo, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except FileNotFoundError:
        return ""

# Configuración básica de logs
logging.basicConfig(
    filename="procesar_facturas.log",  # Archivo donde se guardan los logs
    level=logging.ERROR,  # Guardar solo errores y más graves
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ---------- Interfaz ----------
def mostrar_login():

    render_login_navbar()

    st.markdown('<div class="recuadro-login">', unsafe_allow_html=True)
    st.title("🔐 Iniciar Sesión")
    st.markdown("### Convertidor pedidos ET\nInicia sesión con tu cuenta de Microsoft para continuar.")

    if st.button("🚀 Iniciar sesión", type="primary"):
        auth_url = iniciar_autenticacion()
        abrir_en_nueva_pestana(auth_url)
        st.info("Se abrió una pestaña. Habilita las pestañas emergentes para esta pantalla si no se abre automáticamente.")
    
    st.markdown('</div>', unsafe_allow_html=True)

    render_footer()

def mostrar_aplicacion():

    render_header()  

    st.markdown('<div class="center-wrap">', unsafe_allow_html=True)
    pdf_file = st.file_uploader(
        "Selecciona un PDF de factura",
        type=["pdf"],
        help="Sube un PDF que contenga tablas con productos y la etiqueta 'TIENDA'."
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if pdf_file is not None:
        with st.spinner("Extrayendo datos, dame unos segundos…"):
            try:
                output_excel, nombre_archivo = procesar_pdf(pdf_file, pdf_file.name)

                bytes_data = output_excel.getvalue()  # guardamos el contenido en memoria
                output_excel.close()    

                st.success("✅ ¡PDF procesado!")

                st.markdown('<div class="center-wrap">', unsafe_allow_html=True)
                st.download_button(
                    "📥 Descargar",
                    data=bytes_data,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                st.markdown('</div>', unsafe_allow_html=True)

            except Exception as e:
                st.error("❌ Error al procesar el PDF.")
                with st.expander("Detalles del error"):
                    st.code(str(e))
                logging.exception(f"[{datetime.now()}] Error procesando archivo: {pdf_file.name}")

                # Vista previa de la tabla contenida en el Excel
            if bytes_data:
                bytes_data.seek(0)  # Asegurarse de que el puntero esté al inicio
                try:
                    df_preview = pd.read_excel(bytes_data)
                    st.markdown('<div class="center-preview">', unsafe_allow_html=True)
                    st.markdown("#### 📋 Vista previa de los datos extraidos")
                    st.dataframe(df_preview, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)       
                except Exception as e:
                    st.warning("⚠️ No se pudo mostrar la vista previa de los datos.")
                    logging.exception(f"[{datetime.now()}] Error mostrando vista previa: {e}")

    render_footer()

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
    # si viene ?logout=1 desde el header
    if st.query_params.get("logout"):
        cerrar_sesion()
        return

if __name__ == "__main__":
    iniciar_cron()
    main()