import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
import requests

def exportar_plantilla(bytes_data: bytes) -> BytesIO:
    
    mapeo_columnas = {
        "Tienda": "Tienda",
        "Código": "Referencia",
        "Cantidad": "Unidades"
    }

    # 1️⃣ Cargar archivo base de plantilla + Abrir con openpyxl (manteniendo macros)
    ruta_plantilla = "SaeGA v2.0.2 - Plantilla - copia para Importador de Pedidos - copia.xlsm"
    with open(ruta_plantilla, "rb") as f:
        archivo_base = BytesIO(f.read())

    wb = load_workbook(archivo_base, keep_vba=True)
    ws = wb["Pedidos"]  # cambiar por hoja destino
    tabla = ws.tables["tblPedidos"]  # cambiar por tabla destino

    # 2️⃣ Leer DataFrame origen
    df_origen = pd.read_excel(BytesIO(bytes_data))

    # 3️⃣ Encabezados de la tabla destino
    ref = tabla.ref
    rango = ws[ref]
    encabezados_tabla = [celda.value for celda in rango[0]]

    # 4️⃣ Preparar fila de ejemplo para copiar estilos
    fila_ejemplo = rango[1]

    # 5️⃣ Ajustar número de filas de la tabla destino
    num_filas_actual = len(rango) - 1  # menos fila de encabezado
    num_filas_necesarias = len(df_origen)
    if num_filas_necesarias > num_filas_actual:
        # Insertar filas necesarias justo debajo de los encabezados
        ws.insert_rows(rango[1][0].row + 1, amount=num_filas_necesarias - num_filas_actual)

    # 6️⃣ Rellenar datos y copiar estilos
    for i, fila_origen in enumerate(df_origen.itertuples(index=False), start=1):
        for col_df, col_tabla in mapeo_columnas.items():
            if col_df in df_origen.columns and col_tabla in encabezados_tabla:
                idx_col = encabezados_tabla.index(col_tabla)
                celda_destino = ws.cell(row=fila_ejemplo[0].row + i, column=idx_col + 1)

                # Rellenar valor
                celda_destino.value = getattr(fila_origen, col_df)

                # Copiar estilos de la fila de ejemplo
                celda_destino.font = copy(fila_ejemplo[idx_col].font)
                celda_destino.border = copy(fila_ejemplo[idx_col].border)
                celda_destino.fill = copy(fila_ejemplo[idx_col].fill)
                celda_destino.number_format = copy(fila_ejemplo[idx_col].number_format)
                celda_destino.alignment = copy(fila_ejemplo[idx_col].alignment)

    # 7️⃣ Guardar a BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def subir_a_sharepoint(bytes_io: BytesIO, nombre_archivo: str, access_token: str) -> bool:
    site_id = "saboraespana.sharepoint.com:/sites/departamento.ti"
    drive_id = "b!Y2iXc4H7m3x8eF0z3K9Jt2v1L6gR5QW8x9y0Z1A2B3C4D5E6F7G8H9I0J1K2L3M4"
    carpeta_path = "General/PoC Plantillas SaEGA"

    try:
        # Asegurarnos de que BytesIO está al inicio
        bytes_io.seek(0)

        # Construir URL de subida
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{carpeta_path}/{nombre_archivo}:/content"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }

        response = requests.put(url, headers=headers, data=bytes_io.read())
        response.raise_for_status()

        return True
    except Exception as e:
        print(f"Error subiendo archivo a SharePoint: {e}")
        return False