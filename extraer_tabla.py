from io import BytesIO
import json
import pdfplumber
import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import time
import requests
from urllib.parse import quote
import streamlit as st
from openpyxl.utils import get_column_letter

# 📋 Columnas de salida
COLUMNAS = ["Tienda",                # Relleno
    "Código",                        # Relleno
    "Descripción de artículo",       # Vacío
    "Cantidad",                      # Relleno
    "Precio por unidad",             # Vacío
    "% de descuento",                # Vacío
    "Precio después del descuento",  # Vacío
    "Indicador de impuestos",        # Vacío
    "Total (ML)",                    # Vacío
    "Unidad de negocio",             # Vacío
    "Código de unidad de medida",    # Vacío
    "Precio de coste Departamento"   # Vacío
]

@st.cache_data(ttl=300)  # Cache for 5 minutes
def obtener_orden_maestro_cached(access_token: str) -> list:
    """Cached version of SharePoint master order retrieval"""
    hostname="saboraespana.sharepoint.com"
    site_name="DepartamentodeProducto"
    file_path="General/Aplicaciones/Cadena de Suministro/Herramienta de Aprovisionamiento v1.0.2.xlsx"
    file_path = quote(file_path)
    nombre_hoja = "SURFACE"
    nombre_tabla = "OrdenPreparacion"
    columna_codigos = "SKU"

    headers = {"Authorization": f"Bearer {access_token}"}

    # Get siteId
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    # Download file
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
    file_resp = requests.get(file_url, headers=headers)
    file_resp.raise_for_status()

    # Load with openpyxl
    wb = load_workbook(filename=BytesIO(file_resp.content), data_only=True)
    ws = wb[nombre_hoja]
    tabla = ws.tables[nombre_tabla]
    ref = tabla.ref
    rango = ws[ref]

    # Convert to DataFrame
    contenido = [[celda.value for celda in fila] for fila in rango]
    df_maestro = pd.DataFrame(contenido[1:], columns=contenido[0])

    # Extract and normalize codes
    orden_maestro = (
        df_maestro[columna_codigos]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lstrip("0")
        .tolist()
    )

    return orden_maestro

def procesar_pdf(file_stream, nombre_pdf, sesion):
    NOMBRE_BASE = os.path.splitext(nombre_pdf)[0]

    # 📥 Extract values from first table - optimized single read
    pdf_content = file_stream.read()
    file_stream.seek(0)  # Reset for later use
    
    with pdfplumber.open(BytesIO(pdf_content)) as pdf:
        primera_pagina = pdf.pages[0]
        tablas = primera_pagina.extract_tables()

        try:
            valores = tablas[0][1]
            valor_pedido = valores[1]
        except Exception as e:
            print("⚠️ Error al extraer datos de la tabla:", e)
            valor_pedido = "PEDIDO_NO_ENCONTRADO"

    def extraer_tabla(pdf_bytes):
        filas_resultado = []
        tienda_detectada = ""

        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            all_text = ""
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    all_text += texto_pagina + "\n"
            
            lineas = all_text.split("\n")
            
            for linea in lineas:
                linea = linea.strip()

                # 1️⃣ Buscar tienda
                for linea in lineas:
                    linea = linea.strip()
                    match_tienda = re.search(r"TIENDA\s+(\d+)", linea.upper())
                    if match_tienda:
                        tienda_detectada = match_tienda.group(1)
                        break # Nos quedamos con el primer valor
                
                # 2️⃣ Buscar líneas que empiezan con código
                for linea in lineas:
                    linea = linea.strip()
                    match_codigo = re.match(r"^(\d+)\s+(.*)", linea)
                    if match_codigo:
                        codigo = match_codigo.group(1)

                    if codigo and uds:
                        filas_temporales.append((codigo, uds))

                        if codigo and uds:
                            fila = [
                                f"PEDIDO PC{valor_pedido} TIENDA {tienda_detectada}",  # Tienda
                                codigo,
                                "", uds, "", "", "", "", "", "001", "", "985"
                            ]
                            filas_resultado.append(fila)

        return filas_resultado   #, tienda_detectada

    def ordenar_lineas(df: pd.DataFrame, orden_maestro: list) -> pd.DataFrame:
        # Asegurarse de trabajar con una copia
        df = df.copy()

        # Convertir la columna 'Código' a tipo Categorical con el orden maestro
        # Los códigos que no estén en la lista se convertirán a NaN por defecto
        df['Código'] = pd.Categorical(df['Código'], categories=orden_maestro, ordered=True)

        # Ordenar el DataFrame. Los valores NaN se ordenan al final por defecto.
        df = df.sort_values('Código', na_position='last')

        # Si se desea, se puede volver a convertir la columna a string
        df['Código'] = df['Código'].astype(str)

        return df

    # ▶️ Execute with optimizations
    filas = extraer_tabla(pdf_content)
    df = pd.DataFrame(filas, columns=COLUMNAS)

    orden_maestro = obtener_orden_maestro_cached(sesion)
    df = ordenar_lineas(df, orden_maestro)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
        
        # Add table formatting
        workbook = writer.book
        worksheet = writer.sheets["Datos"]
        
        ultima_fila = len(df) + 1
        ultima_columna = len(COLUMNAS)
        letra_ultima_columna = get_column_letter(ultima_columna)
        rango_tabla = f"A1:{letra_ultima_columna}{ultima_fila}"

        tabla = Table(displayName="TablaDatos", ref=rango_tabla)
        estilo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        tabla.tableStyleInfo = estilo
        worksheet.add_table(tabla)

    output.seek(0)
    nombre_final = f"Factura_{NOMBRE_BASE}.xlsx".replace(" ", "_")
    return output, nombre_final
