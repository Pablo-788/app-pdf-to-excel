import base64
import streamlit as st
import logging
from auth import iniciar_autenticacion, cerrar_sesion
from extraer_tabla import procesar_pdf
from datetime import datetime
import pandas as pd
import time
from io import BytesIO
from exportacion_plantilla import (
    limpiar_entradas_mgraph,
    escribir_datos_mgraph,
    copy_file_on_sharepoint,
    descargar_file_de_sharepoint
)

def init_state():
    defaults = {
        "last_pdf_key": None,            # identifica si el PDF es nuevo
        "export_done": False,            # ya se volcó a Excel este PDF
        "excel_final_bytes": None,       # bytes del xlsm actualizado para descargar
        "uploaded_to_sharepoint": False  # ya se subió a SharePoint
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


RUTA_PLANTILLA = "SaeGA v2.0.2 - Plantilla - copia para Importador de Pedidos - copia.xlsm"

APP_TITLE   = "Convertidor Pedidos ET → Excel"
APP_VERSION = "0.3.16"

PRIMARY_COLOR     = "#0d6efd"
BRAND_BLUE_LIGHT  = "#E6F0FA"
BRAND_BLUE_DARK   = "#003366"
LOGO_PATH = "LOGO_SAE.png"

def inject_styles():
    st.markdown(f"""
    <style>
    /* ======= NAVBAR ======= */
    .navbar {{ display:flex; justify-content:space-between; align-items:center; 
               background:{BRAND_BLUE_LIGHT}; padding:10px 24px; }}
    .navbar-left {{ display:flex; align-items:center; gap:12px; }}
    .navbar-left img {{ height:40px; }}
    .navbar-title {{ font-weight:600; color:{BRAND_BLUE_DARK}; font-size:18px; }}
    .navbar-right {{ display:flex; align-items:center; gap:12px;justify-content: flex-end;}}

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
        background-color: #bc9a5f;
        border-radius: 16px;
        padding: 40px;
        max-width: 500px;
        margin: 50px auto;
        text-align: center;
        color: #fff;
    }}

    .recuadro-login h1 {{
        font-size: 32px;
        margin-bottom: 16px;
    }}

    .recuadro-login h3 {{
        font-size: 20px;
        margin-bottom: 24px;
    }}

    /* ======= BOTONES ======= */
    .stButton > button,
    .stDownloadButton > button,
    .stDownloadButton > a {{
        background:{PRIMARY_COLOR} !important;
        color:#fff !important; 
        border: 0 !important;
        border-radius: 8px !important;
        padding: 10px 16px !important;
        font-size: 16px !important;

        display: inline-flex !important;
        justify-content: center !important;
        align-items: center !important;

        max-width: 250px !important;
        width: auto !important;
        margin: 0.5rem 0 !important;
    }}

    .stButton > button:hover,
    .stDownloadButton > button:hover,
    .stDownloadButton > a:hover {{
        background: #0056b3 !important;  /* cambia este color si quieres otro hover */
    }}

    /* ======= FOOTER ======= */
    .footer {{ text-align:center !important; margin:48px 0 16px !important; color:#111 !important; }}
    .footer .brand {{ color:{PRIMARY_COLOR} !important; font-weight:700 !important; }}
    
    </style>
    """, unsafe_allow_html=True)

def cargar_logo_base64(ruta_logo: str = LOGO_PATH) -> str:
    """Carga el logo desde archivo y devuelve su contenido en base64. Si falla, devuelve cadena vacía."""
    try:
        with open(ruta_logo, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except FileNotFoundError:
        return ""

def render_header():
    inject_styles()

    logo_b64 = cargar_logo_base64()
    info = st.session_state.get("user_info", {})
    name  = info.get("name", "Usuario")
    email = info.get("preferred_username", "usuario@example.com")

    st.markdown(f"""
    <style>
        /* Contenedor general del header */
        .navbar-container {{
            background: {BRAND_BLUE_LIGHT};
            padding: 10px 24px;
            border-radius: 0;
        }}
    </style>
    """, unsafe_allow_html=True)

    # --- Usamos columnas para maquetar ---
    with st.container():
        col1, col2 = st.columns([6, 2])  # proporción izquierda/derecha
        with col1:
            st.markdown(
                f"""
                <div class="navbar-container navbar-left">
                    {f'<img src="data:image/png;base64,{logo_b64}" alt="logo" />' if logo_b64 else ''}
                    <span class="navbar-title">{APP_TITLE}</span>
                </div>
                """,
                unsafe_allow_html=True
            )
        with col2:
            st.markdown('<div class="navbar-container navbar-right">', unsafe_allow_html=True)

            # ---------- Menú ----------
            with st.popover(name, use_container_width=True):
                st.markdown(f"""
                    <p><b>Nombre:</b> {name}</p>
                    <p><b>Email:</b> {email}</p>
                    <hr>
                """, unsafe_allow_html=True)
                # El botón ahora llama directamente a la función de cierre de sesión
                if st.button("Cerrar sesión"):
                    cerrar_sesion()
                st.markdown(f"""
                    <p style="margin-top:8px; font-size:12px; opacity:.7;">
                        Versión: {APP_VERSION}
                    </p>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

def render_footer():
    inject_styles()
    st.markdown(f"""
    <div class="footer">
        Made with sweetness by <span class="brand">SaE Tech Team!</span><br/>
        Version: {APP_VERSION}
    </div>
    """, unsafe_allow_html=True)

def render_login_navbar():
    inject_styles()
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

# ---------- ---------- Configuración básica de logs ---------- ----------
logging.basicConfig(
    filename="procesar_facturas.log",  # Archivo donde se guardan los logs
    level=logging.ERROR,  # Guardar solo errores y más graves
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ---------- ---------- ---------- Interfaz ---------- ---------- ----------
def mostrar_login():
    inject_styles()
    render_login_navbar()
    st.markdown(f"""
    <div class="recuadro-login">
        <h1>🔐 Iniciar Sesión</h1>
        <h3>Convertidor pedidos ET</h3>
        <p>Inicia sesión con tu cuenta de Microsoft para continuar.</p>
    </div>
    """, unsafe_allow_html=True)
    _, col2, _ = st.columns([1, 1.5, 1])
    with col2:
        auth_url = iniciar_autenticacion()
        st.markdown(
            f"""
            <div style="text-align: center; margin-top: 20px;">
                <a href="{auth_url}" target="_self" class="btn btn-primary" 
                style="background-color: #0d6efd; border: none; color: white; 
                        padding: 10px 20px; text-align: center; text-decoration: none; 
                        display: inline-block; font-size: 16px; border-radius: 6px; 
                        font-weight: 500; cursor: pointer;">
                    🚀 Iniciar sesión
                </a>
            </div>
            """,
            unsafe_allow_html=True)

    render_footer()

def mostrar_aplicacion():
    inject_styles()
    init_state()
    render_header()

    st.markdown('<div class="center-wrap">', unsafe_allow_html=True)

    # Paso 1: Subida de PDF
    pdf_file = st.file_uploader(
        "📄 Selecciona un PDF de factura",
        type=["pdf"],
        help="Sube un PDF que contenga tablas con productos y la etiqueta 'TIENDA'."
    )

    # Detectar cambio de PDF para resetear estado
    if pdf_file is not None:
        pdf_key = f"{pdf_file.name}-{getattr(pdf_file, 'size', 0)}"
        if st.session_state.last_pdf_key != pdf_key:
            st.session_state.last_pdf_key = pdf_key
            st.session_state.export_done = False
            st.session_state.uploaded_to_sharepoint = False
            st.session_state.excel_final_bytes = None
            st.session_state.final_file_name = None
            st.session_state.download_ready = False

    st.markdown('</div>', unsafe_allow_html=True)

    # Si ya tenemos el archivo final generado, mostramos botón descarga
    if st.session_state.get('download_ready'):
        st.success("🎉 ¡Proceso completado correctamente!")
        _, col_center, _ = st.columns([1, 2, 1])
        with col_center:
            st.download_button(
                label="⬇️ Descargar plantilla modificada",
                data=st.session_state.excel_final_bytes,
                file_name=st.session_state.final_file_name,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                key='final_download_button'
            )
        st.info("Para procesar otro archivo, sube un nuevo PDF.")
        return

    # Si hay PDF, procesamos
    if pdf_file is not None:
        with st.spinner("📥 Extrayendo datos del PDF…"):
            try:
                output_excel, nombre_archivo_base = procesar_pdf(
                    pdf_file, pdf_file.name, st.session_state.access_token
                )
                output_excel.seek(0)
                bytes_data = output_excel.getvalue()
                output_excel.close()
                st.success("✅ Datos extraídos correctamente")
            except Exception as e:
                st.error("❌ Error al procesar el PDF.")
                with st.expander("Detalles del error"):
                    st.code(str(e))
                logging.exception(f"[{datetime.now()}] Error procesando archivo: {pdf_file.name}")
                return

        # Vista previa
        try:
            df_preview = pd.read_excel(BytesIO(bytes_data))
            st.markdown("#### 📋 Vista previa de los datos extraídos")
            st.dataframe(df_preview, use_container_width=True)
        except Exception as e:
            st.warning("⚠️ No se pudo mostrar la vista previa.")
            logging.exception(f"[{datetime.now()}] Error mostrando vista previa: {e}")

        # Paso 2: número de usuario
        numero_usuario = st.text_input(
            "Introduce un número de 6 dígitos para nombrar la plantilla de SaeGA",
            max_chars=6
        )

        if numero_usuario and numero_usuario.isdigit() and len(numero_usuario) == 6:
            with st.spinner("⚙️ Procesando datos en la plantilla de SharePoint…"):
                try:
                    SHAREPOINT_DRIVE_ID = "b!u5eMfc6yEEyNixso-TGvPvlFaVzwwS9Kg8-2acKgkFgmw3vFAD1wS7AcibxLgDUS"
                    SHAREPOINT_TEMPLATE_ITEM_ID = "01TGJYFESV66I2HZQYCZCIDYHYPLIT6HZQ"

                    limpiar_entradas_mgraph(
                        access_token=st.session_state.access_token,
                        drive_id=SHAREPOINT_DRIVE_ID,
                        item_id=SHAREPOINT_TEMPLATE_ITEM_ID,
                        nombre_tabla="tblPedidos"
                    )
                    escribir_datos_mgraph(
                        access_token=st.session_state.access_token,
                        drive_id=SHAREPOINT_DRIVE_ID,
                        item_id=SHAREPOINT_TEMPLATE_ITEM_ID,
                        bytes_data=bytes_data
                    )
                    copy_file_on_sharepoint(
                        access_token=st.session_state.access_token,
                        drive_id=SHAREPOINT_DRIVE_ID,
                        item_id=SHAREPOINT_TEMPLATE_ITEM_ID,
                        new_file_name=f"{numero_usuario} - SaeGA.xlsm"
                    )

                    time.sleep(3)  # Espera para que SharePoint sincronice

                    downloaded_bytes = descargar_file_de_sharepoint(
                        access_token=st.session_state.access_token,
                        drive_id=SHAREPOINT_DRIVE_ID,
                        item_id=SHAREPOINT_TEMPLATE_ITEM_ID
                    )

                    if downloaded_bytes:
                        st.session_state.excel_final_bytes = downloaded_bytes
                        st.session_state.final_file_name = f"{numero_usuario} - SaeGA.xlsm"
                        st.session_state.download_ready = True
                        st.rerun()
                    else:
                        st.error("❌ No se pudo descargar el archivo actualizado.")
                except Exception as e:
                    st.error("❌ Ocurrió un error en el proceso.")
                    with st.expander("Detalles del error"):
                        st.code(str(e))

        elif numero_usuario:
            st.warning("⚠️ Debes introducir exactamente 6 números.")

    render_footer()
