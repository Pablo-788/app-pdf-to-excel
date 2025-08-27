from io import BytesIO
import pdfplumber
import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# 📋 Columnas de salida
COLUMNAS = ["Número de artículo",    # Relleno
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

def procesar_pdf(file_stream, nombre_pdf, session):
    NOMBRE_BASE = os.path.splitext(nombre_pdf)[0]

    # 📥 Extraer valores desde la primera tabla real del PDF
    with pdfplumber.open(file_stream) as pdf:
        primera_pagina = pdf.pages[0]
        tablas = primera_pagina.extract_tables()

        try:
            # Esta función ahora podría usar la 'session' si fuera necesario
            def obtener_orden_maestro(pdf_tables, session_info):
                # Aquí iría la lógica que necesite la sesión
                # Por ahora, extraemos de las tablas como antes
                return pdf_tables[0][1]

            valores = obtener_orden_maestro(tablas, session)
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
    file_stream.seek(0) # Reiniciamos el puntero del stream por si se ha movido
    filas, tienda_detectada = extraer_tabla(file_stream)
    df = pd.DataFrame(filas, columns=COLUMNAS)

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

    # --- Bloque de código corregido y activado ---
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
    # --- Fin del bloque ---

    nombre_final = f"Factura_{NOMBRE_BASE}.xlsx".replace(" ", "_")

    return nuevo_output, nombre_final