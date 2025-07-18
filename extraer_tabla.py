from io import BytesIO
import pdfplumber
import pandas as pd
import os
import re
from openpyxl import load_workbook

# 📋 Columnas de salida
COLUMNAS = ["Número de artículo",             # Relleno
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

def procesar_pdf(file_stream, nombre_pdf):
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
                                codigo,      # Número de artículo
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

    # ▶️ Ejecutar
    filas, tienda_detectada = extraer_tabla(file_stream)
    df = pd.DataFrame(filas, columns=COLUMNAS)

    # 🧾 Guardar Excel en memoria
    output = BytesIO()
    df.to_excel(output, index=False, sheet_name="Datos")
    output.seek(0)

    # Reabrimos el archivo en memoria
    wb = load_workbook(output)
    ws = wb["Datos"]

    # Escribimos el resumen en una celda fuera de la tabla
    ultima_fila = ws.max_row
    resumen_texto = f"PEDIDO PC{valor_pedido} TIENDA {tienda_detectada}"
    ws.cell(row=ultima_fila + 2, column=1).value = resumen_texto

    # Guardamos los cambios en un nuevo BytesIO
    nuevo_output = BytesIO()
    wb.save(nuevo_output)
    nuevo_output.seek(0)

    nombre_final = f"Factura_{NOMBRE_BASE}.xlsx".replace(" ", "_")

    return nuevo_output, nombre_final