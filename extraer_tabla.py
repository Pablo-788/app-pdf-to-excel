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
        dentro_de_tabla = False
        indices_columnas = {}

        with pdfplumber.open(pdf_stream) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto:
                    continue
                lineas = texto.split("\n")

                for linea in lineas:
                    linea = linea.strip()
                    
                    # 🏪 Detectar tienda
                    match_tienda = re.search(r"TIENDA\s+(\d+)", linea.upper())
                    if match_tienda:
                        tienda_detectada = match_tienda.group(1)

                    # 🧱 Detectar cabecera de la "tabla"
                    if not dentro_de_tabla and re.search(r"\bCÓDIGO\b.*\bUDS\.\b.*\bIMPORTE\b", linea.upper()):
                        headers = linea.upper().split()
                        for i, col in enumerate(headers):
                            if col in ["CÓDIGO", "UDS."]:
                                indices_columnas[col] = i
                        dentro_de_tabla = True
                        continue

                    # 📦 Procesar línea con producto
                    if dentro_de_tabla:
                        partes = linea.split()
                        if len(partes) < max(indices_columnas.values()) + 1:
                            continue  # Línea incompleta

                        try:
                            codigo = partes[indices_columnas["CÓDIGO"]]
                            uds = partes[indices_columnas["UDS."]]

                            fila = [
                                codigo,  # Número de artículo
                                "",      # Descripción
                                uds,     # Cantidad
                                "", "", "", "", "", "", "", ""  # Vacíos
                            ]
                            filas_resultado.append(fila)
                        except Exception as e:
                            print(f"❌ Error al procesar línea: {linea} -> {e}")
                            continue

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
    resumen_texto = f"PEDIDO PC{valor_pedido} TIENDA {tienda_detectada}"
    ws["B2"] = resumen_texto

    # Guardamos los cambios en un nuevo BytesIO
    nuevo_output = BytesIO()
    wb.save(nuevo_output)
    nuevo_output.seek(0)

    nombre_final = f"Factura_{NOMBRE_BASE}.xlsx".replace(" ", "_")

    return nuevo_output, nombre_final