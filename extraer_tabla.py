from io import BytesIO
import json
import pdfplumber
import pandas as pd
import os
import re
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
    "Unidad de negocio",             # Fijo "001"
    "Código de unidad de medida",    # Vacío
    "Precio de coste Departamento"   # Fijo "985"
]

@st.cache_data(ttl=300)  # El caché sigue siendo útil para evitar llamadas repetidas a la API
def obtener_orden_maestro_cached(access_token: str) -> list:
    """
    Versión optimizada que lee solo la columna necesaria de la tabla en SharePoint
    usando la API de Microsoft Graph, sin descargar el archivo Excel completo.
    """
    hostname = "saboraespana.sharepoint.com"
    site_name = "DepartamentodeProducto"
    file_path = "General/Aplicaciones/Cadena de Suministro/Herramienta de Aprovisionamiento v1.0.2.xlsx"
    nombre_hoja = "SURFACE"
    nombre_tabla = "OrdenPreparacion"
    columna_codigos = "SKU"

    headers = {"Authorization": f"Bearer {access_token}"}
    
    # URL base para las llamadas a la API de Graph
    graph_url = "https://graph.microsoft.com/v1.0"

    try:
        # 1. Obtener siteId
        site_url = f"{graph_url}/sites/{hostname}:/sites/{site_name}"
        site_resp = requests.get(site_url, headers=headers)
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]

        # 2. Obtener driveItemId del archivo
        # La ruta debe estar codificada para la URL, pero sin codificar las barras '/'
        file_path_encoded = quote(file_path, safe='')
        item_url = f"{graph_url}/sites/{site_id}/drive/root:/{file_path_encoded}"
        item_resp = requests.get(item_url, headers=headers)
        item_resp.raise_for_status()
        item_id = item_resp.json()["id"]

        # 3. Leer directamente el rango de la columna de la tabla
        # Esta es la llamada clave que evita la descarga del archivo
        column_data_url = (
            f"{graph_url}/sites/{site_id}/drive/items/{item_id}/workbook/tables('{nombre_tabla}')"
            f"/columns('{columna_codigos}')/range"
        )
        
        # Usamos $select para pedir solo el campo 'values' y reducir la respuesta
        params = {"$select": "values"}
        data_resp = requests.get(column_data_url, headers=headers, params=params)
        data_resp.raise_for_status()
        
        # El resultado es un JSON con una matriz de valores
        # [['SKU'], ['12345'], ['67890'], [''], ...]
        json_data = data_resp.json()
        values = json_data.get("values", [])

        # 4. Procesar la lista de códigos directamente desde el JSON
        if not values or len(values) < 2:  # Si no hay datos o solo la cabecera
            return []

        # Omitimos la primera fila (cabecera) y procesamos el resto
        orden_maestro = [
            str(row[0]).strip().lstrip("0")
            for row in values[1:]
            if row and row[0] is not None and str(row[0]).strip()
        ]

        return orden_maestro

    except requests.exceptions.RequestException as e:
        # Manejo de errores de red o de la API
        st.error(f"Error al contactar con la API de Microsoft Graph: {e}")
        return []
    except (KeyError, IndexError) as e:
        # Manejo de errores por respuesta inesperada del JSON
        st.error(f"Error al procesar la respuesta de la API (estructura inesperada): {e}")
        return []

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
        filas_temporales = []
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
                match_tienda = re.search(r"TIENDA\s+(\d+)", linea.upper())
                if match_tienda:
                    tienda_detectada = match_tienda.group(1)
                    
                    # Añadir todas las filas temporales con la tienda detectada
                    for codigo, uds in filas_temporales:
                        fila = [
                            f"PEDIDO PC{valor_pedido} TIENDA {tienda_detectada}",
                            codigo,
                            "", uds, "", "", "", "", "", "001", "", "985"
                        ]
                        filas_resultado.append(fila)
                    filas_temporales.clear()
                    continue

                # 2️⃣ Buscar líneas que empiezan con código
                match_codigo = re.match(r"^(\d+)\s+(.*)", linea)
                if match_codigo:
                    codigo = match_codigo.group(1)

                    # 3️⃣ Buscar unidades tipo "6,000" en la línea
                    partes = linea.split()
                    uds = next((p for p in partes if re.match(r"^\d+,\d{3}$", p)), None)

                    if codigo and uds:
                        filas_temporales.append((codigo, uds))

        return filas_resultado

    def ordenar_lineas(df, orden_maestro):
        pos = {codigo: i for i, codigo in enumerate(orden_maestro)}
        df["orden_idx"] = df["Código"].map(pos).fillna(float("inf"))
        df = df.sort_values("orden_idx").drop(columns=["orden_idx"])
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
