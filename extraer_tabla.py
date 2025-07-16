from io import BytesIO
import pdfplumber
import pandas as pd
import os
import re

# 📋 Columnas de salida
COLUMNAS = ["CÓDIGO", "DESCRIPCIÓN", "CAJAS", "UDS.", "PRECIO", "IMPORTE", "TIENDA", "PEDIDO", "FECHA"]

def procesar_pdf(file_stream, nombre_pdf):
    NOMBRE_BASE = os.path.splitext(nombre_pdf)[0]

    # 📥 Extraer valores desde la primera tabla real del PDF
    with pdfplumber.open(file_stream) as pdf:
        primera_pagina = pdf.pages[0]
        tablas = primera_pagina.extract_tables()

        try:
            valores = tablas[0][1]  # Fila de datos (segunda fila)
            valor_fecha = valores[0].replace("/", "-")
            valor_pedido = valores[1]
        except Exception as e:
            print("⚠️ Error al extraer datos de la tabla:", e)
            valor_fecha = "FECHA_NO_ENCONTRADA"
            valor_pedido = "PEDIDO_NO_ENCONTRADO"

    def extraer_tabla(pdf_stream):
        filas_sin_tienda = []
        tienda_detectada = ""

        with pdfplumber.open(pdf_stream) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto:
                    continue
                lineas = texto.split("\n")

                for linea in lineas:
                    linea = linea.strip()

                    match_tienda = re.search(r"TIENDA\s+(\d+)", linea.upper())
                    if match_tienda:
                        tienda_detectada = match_tienda.group(1)

                    match_codigo = re.match(r"^(\d+)\s+(.*)", linea)
                    if match_codigo:
                        partes = linea.split()
                        if len(partes) < 6:
                            continue

                        try:
                            importe = partes[-1]
                            precio = partes[-2]
                            uds = partes[-3]
                            cajas = partes[-4]
                            codigo = partes[0]
                            descripcion = " ".join(partes[1:-4])

                            fila = [codigo, descripcion, cajas, uds, precio, importe, None, None, None]
                            filas_sin_tienda.append(fila)
                        except Exception as e:
                            print(f"❌ Error en línea: {linea} -> {e}")
                            continue

        for fila in filas_sin_tienda:
            fila[-1] = valor_fecha
            fila[-2] = valor_pedido
            fila[-3] = tienda_detectada

        return filas_sin_tienda

    # ▶️ Ejecutar
    filas = extraer_tabla(file_stream)
    df = pd.DataFrame(filas, columns=COLUMNAS)

    # 🧠 Conversión de tipos
    df["CÓDIGO"] = pd.to_numeric(df["CÓDIGO"], errors="coerce").astype("Int64")
    df["CAJAS"] = pd.to_numeric(df["CAJAS"], errors="coerce").astype("Int64")
    df["TIENDA"] = pd.to_numeric(df["TIENDA"], errors="coerce").astype("Int64")
    df["PEDIDO"] = pd.to_numeric(df["PEDIDO"], errors="coerce").astype("Int64")

    # Sustituir comas por puntos antes de convertir a float
    df["UDS."] = df["UDS."].str.replace(",", ".", regex=False)
    df["PRECIO"] = df["PRECIO"].str.replace(",", ".", regex=False)
    df["IMPORTE"] = df["IMPORTE"].str.replace(",", ".", regex=False)

    df["UDS."] = pd.to_numeric(df["UDS."], errors="coerce")
    df["PRECIO"] = pd.to_numeric(df["PRECIO"], errors="coerce")
    df["IMPORTE"] = pd.to_numeric(df["IMPORTE"], errors="coerce")

    df["FECHA"] = pd.to_datetime(df["FECHA"], format="%d-%m-%y", errors="coerce")
    df = df.fillna("")

    # 🧾 Guardar Excel en memoria
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    nombre_final = f"Factura_{NOMBRE_BASE}.xlsx".replace(" ", "_")

    return output, nombre_final