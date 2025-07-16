import streamlit as st
from extraer_tabla import procesar_pdf

st.set_page_config(page_title="Convertidor PDF → Excel", layout="centered")
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