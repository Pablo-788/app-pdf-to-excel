from io import BytesIO
import pdfplumber
import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import time
import requests
from urllib.parse import quote

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

def procesar_pdf(file_stream, nombre_pdf, sesion):
    NOMBRE_BASE = os.path.splitext(nombre_pdf)[0]

    # 📥 Extraer valores desde la primera tabla real del PDF
    with pdfplumber.open(file_stream) as pdf:
        primera_pagina = pdf.pages[0]
        tablas = primera_pagina.extract_tables()

        try:
            valores = tablas[0][1]  # Fila de datos (segunda fila)
            valor_pedido = valores[1]
        except Exception as e:
            print("⚠️ Error al extraer datos de la tabla:", e)
            valor_pedido = "PEDIDO_NO_ENCONTRADO"

    def extraer_tabla(pdf_stream):
        filas_resultado = []
        tienda_detectada = ""

        with pdfplumber.open(pdf_stream) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if not texto_pagina:
                    continue
                lineas = texto_pagina.split("\n")

                for linea in lineas:
                    linea = linea.strip()

                    # 1️⃣ Buscar tienda
                    match_tienda = re.search(r"TIENDA\s+(\d+)", linea.upper())
                    if match_tienda:
                        tienda_detectada = match_tienda.group(1)

                    # 2️⃣ Buscar líneas que empiezan con código
                    match_codigo = re.match(r"^(\d+)\s+(.*)", linea)
                    if match_codigo:
                        codigo = match_codigo.group(1)

                        # 3️⃣ Buscar unidades tipo "6,000" en la línea
                        partes = linea.split()
                        uds = next((p for p in partes if re.match(r"^\d+,\d{3}$", p)), None)

                        if codigo and uds:
                            fila = [
                                f"PEDIDO PC{valor_pedido} TIENDA {tienda_detectada}",  # Tienda
                                codigo,      # Código
                                "",          # Descripción
                                uds,         # Cantidad
                                "",          # Precio por unidad
                                "",          # % de descuento
                                "",          # Precio después del descuento
                                "",          # Indicador de impuestos
                                "",          # Total (ML)
                                "001",       # Unidad de negocio
                                "",          # Código de unidad de medida
                                "985"        # Precio de coste Departamento
                            ]
                            filas_resultado.append(fila)

        return filas_resultado, tienda_detectada

    def ordenar_lineas(df, orden_maestro):
        # Crear diccionario: código → posición en el orden maestro
        pos = {codigo: i for i, codigo in enumerate(orden_maestro)}

        # Columna auxiliar con la posición, por defecto inf si no está en el maestro
        df["orden_idx"] = df["Número de artículo"].map(lambda x: pos.get(x, float("inf")))

        # Ordenar según esa columna y eliminarla
        df = df.sort_values("orden_idx").drop(columns=["orden_idx"])

        return df
    
    # Variable de caché a nivel de función
    _cache_orden_maestro = {
        "timestamp": 0,
        "orden_maestro": []
    }

    def obtener_orden_maestro(access_token, cache_tiempo_seg=5):
        hostname="saboraespana.sharepoint.com"
        site_name="DepartamentodeProducto"
        file_path="Documentos Compartidos/General/Aplicaciones/Cadena de Suministro/Herramienta de Aprovisionamiento v1.0.2.xlsx"
        file_path = quote(file_path)
        nombre_hoja = "SURFACE"
        nombre_tabla = "OrdenPreparacion"
        columna_codigos = "SKU"

        nonlocal _cache_orden_maestro
        ahora = time.time()

        # 1️⃣ Revisar si la cache sigue vigente
        if ahora - _cache_orden_maestro["timestamp"] < cache_tiempo_seg:
            return _cache_orden_maestro["orden_maestro"]

        # 2️⃣ Descargar el archivo de SharePoint en memoria
        """
        Descarga un archivo Excel desde SharePoint y lo devuelve como objeto openpyxl.Workbook.

        :param access_token: Token OAuth2 obtenido con MSAL
        :param hostname: dominio del SharePoint (ej: 'contoso.sharepoint.com')
        :param site_name: nombre del sitio (ej: 'MiProyecto')
        :param file_path: ruta al archivo dentro de la biblioteca (ej: 'Carpeta/archivo.xlsx')
        :return: objeto openpyxl.Workbook
        """
        headers = {"Authorization": f"Bearer {access_token}"}

        # 2️⃣.1 Obtener siteId
        site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
        site_resp = requests.get(site_url, headers=headers)
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]            #saboraespana.sharepoint.com,6764a04e-2820-49f7-87c6-460bb716d51b,1e8bc8aa-29db-463b-958a-9564f0e2b951

#        drives_resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives", headers=headers)
#        print(drives_resp.json())
#        drive_id = drives_resp.json()["id"]          #b!TqBkZyAo90mHxkYLtxbVG6rIix7bKTtGlYqVZPDiuVGl0FEjTV5wQJJdSf4OmGWZ
#        folder_resp = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/General/Aplicaciones/Cadena%20de%20Suministro:/children", headers=headers)
#        print(folder_resp.json())

        # 2️⃣.2 Obtener metadata del archivo (para conseguir itemId)
        file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}"
        file_resp = requests.get(file_url, headers=headers)
        file_resp.raise_for_status()
        item_id = file_resp.json()["id"]

        # 2️⃣.3 Descargar contenido del archivo
        download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
        download_resp = requests.get(download_url, headers=headers)
        download_resp.raise_for_status()

        # 2️⃣.4 Cargarlo en memoria con openpyxl
        wb = load_workbook(filename=BytesIO(download_resp.content), data_only=True)

        # 3️⃣ Leer Excel con openpyxl
        #wb = load_workbook(file_bytes, data_only=True)
        ws = wb[nombre_hoja]

        # 4️⃣ Obtener la tabla por nombre
        tabla = ws.tables[nombre_tabla]
        ref = tabla.ref  # Ejemplo: "A1:C200"
        rango = ws[ref]

        # 5️⃣ Convertir a DataFrame con pandas (igual que antes, pero solo ese rango)
        contenido = [[celda.value for celda in fila] for fila in rango]
        df_maestro = pd.DataFrame(contenido[1:], columns=contenido[0])

        # 6️⃣ Extraer columna de códigos y normalizar
        orden_maestro = (
            df_maestro[columna_codigos]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lstrip("0")
            .tolist()
        )

        # 7️⃣ Actualizar cache
        _cache_orden_maestro["timestamp"] = ahora
        _cache_orden_maestro["orden_maestro"] = orden_maestro

        return orden_maestro

    # ▶️ Ejecutar
    filas, tienda_detectada = extraer_tabla(file_stream)
    df = pd.DataFrame(filas, columns=COLUMNAS)

    # Aquí obtenemos el orden maestro desde SharePoint
    orden_maestro = obtener_orden_maestro(sesion)

    # Aquí llamas a ordenar_lineas
    df = ordenar_lineas(df, orden_maestro)

    # 🧾 Guardar Excel en memoria
    output = BytesIO()
    df.to_excel(output, index=False, sheet_name="Datos")
    output.seek(0)

    wb_tabla = load_workbook(output)
    ws_tabla = wb_tabla["Datos"]

    # Definir rango de la tabla
    ultima_fila = ws_tabla.max_row
    ultima_columna = ws_tabla.max_column
    letra_ultima_columna = ws_tabla.cell(row=1, column=ultima_columna).column_letter
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
    ws_tabla.add_table(tabla)

    # Guardamos el Excel con tabla en memoria
    output_tabla = BytesIO()
    wb_tabla.save(output_tabla)
    output_tabla.seek(0)

    '''
    # Reabrimos el archivo en memoria
    wb = load_workbook(output_tabla)
    ws = wb["Datos"]

    # Escribimos el resumen en una celda fuera de la tabla
    ultima_fila = ws.max_row
    resumen_texto = f"PEDIDO PC{valor_pedido} TIENDA {tienda_detectada}"
    ws.cell(row=ultima_fila + 2, column=1).value = resumen_texto

    # Guardamos los cambios en un nuevo BytesIO
    nuevo_output = BytesIO()
    wb.save(nuevo_output)
    nuevo_output.seek(0)
    '''

    nombre_final = f"Factura_{NOMBRE_BASE}.xlsx".replace(" ", "_")

    return output_tabla, nombre_final