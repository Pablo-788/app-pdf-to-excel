# exportacion_plantilla.py
from io import BytesIO
import requests
from urllib.parse import quote
import os
import pandas as pd
# COM de Excel
import xlwings as xw


# ---- Ajusta esto si quieres un path por defecto para tu .xlsm local ----
RUTA_PLANTILLA_POR_DEFECTO = "SaeGA v2.0.2 - Plantilla - copia para Importador de Pedidos - copia.xlsm"

def limpiar_entradas_xlwings(
    ruta_excel: str,
    hoja: str = "Pedidos",
    nombre_tabla: str = "tblPedidos",
    fila_inicio = 3,
    col_inicio: str = "A",
    col_fin: str = "C",
    ajustar_filas: bool = False
) -> None:
    """
    Limpia las columnas de entrada (A, B, C) usando xlwings.
    Opcionalmente, reduce las filas de la tabla a 1.
    """
    # El bloque 'with' gestiona la apertura y cierre de Excel automáticamente
    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.enable_events = False
        
        wb = app.books.open(ruta_excel)
        ws = wb.sheets[hoja]

        last_row = fila_inicio - 1
        tbl = None
        
        try:
            tbl = ws.tables[nombre_tabla]
            # Si la tabla tiene cuerpo de datos, calcula la última fila
            if tbl.data_body_range is not None:
                existing_rows = tbl.data_body_range.rows.count
                if existing_rows > 0:
                    last_row = tbl.header_row_range.row + existing_rows
        except KeyError:
            # Si no hay tabla, busca la última fila usada en la columna C
            last_row = ws.range(f'{col_fin}{ws.cells.last_cell.row}').end('up').row

        # Limpia el contenido del rango en una sola operación
        if last_row >= fila_inicio:
            ws.range(f'{col_inicio}{fila_inicio}:{col_fin}{last_row}').clear_contents()

        # Opcional: reducir filas de la tabla a 1
        if ajustar_filas and tbl and tbl.data_body_range is not None:
            existing_rows = tbl.data_body_range.rows.count
            # Borra filas de abajo hacia arriba para evitar problemas de índice
            while existing_rows > 1:
                tbl.api.ListRows(existing_rows).Delete()
                existing_rows -= 1
        
        wb.save()
        wb.close()


def exportar_directo_excel_xlwings(
    ruta_excel: str,
    bytes_data: bytes,
    hoja: str = "Pedidos",
    nombre_tabla: str = "tblPedidos",
    celda_inicio: str = "A3",
    columnas_df: tuple = ("Tienda", "Código", "Cantidad")
) -> None:
    """
    Escribe datos de un DataFrame en un Excel usando xlwings.
    Ajusta el tamaño de la tabla de destino para que coincida con los datos.
    """
    df = pd.read_excel(BytesIO(bytes_data))
    if df.empty:
        return
        
    # Asegúrate de que las columnas existen y están en el orden correcto
    try:
        df_to_write = df[list(columnas_df)]
    except KeyError as e:
        raise ValueError(f"Falta una columna requerida en el DataFrame: {e}")

    n_nuevas_filas = len(df_to_write)

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.enable_events = False

        wb = app.books.open(ruta_excel)
        ws = wb.sheets[hoja]

        # ===== AJUSTE DE LA TABLA (SI EXISTE) =====
        try:
            tbl = ws.tables[nombre_tabla]
            existing_rows = 0
            if tbl.data_body_range is not None:
                existing_rows = tbl.data_body_range.rows.count

            # Añade o borra filas para que coincida con el DataFrame
            if existing_rows > n_nuevas_filas:
                # Borra filas sobrantes (de abajo hacia arriba)
                for i in range(existing_rows, n_nuevas_filas, -1):
                    tbl.api.ListRows(i).Delete()
            elif existing_rows < n_nuevas_filas:
                # Añade las filas que faltan
                for _ in range(n_nuevas_filas - existing_rows):
                    tbl.api.ListRows.Add()
        except KeyError:
            # Si la tabla no existe, no hacemos nada y simplemente escribimos en el rango
            pass

        # ===== ESCRITURA DIRECTA Y EFICIENTE DEL DATAFRAME =====
        # Escribe todo el DataFrame en una sola operación, sin índice ni cabecera
        ws.range(celda_inicio).options(index=False, header=False).value = df_to_write
        
        wb.save()
        wb.close()

def subir_a_sharepoint(
    bytes_io: BytesIO,
    nombre_archivo: str,
    access_token: str,
    hostname: str = "saboraespana.sharepoint.com",
    site_name: str = "departamento.ti",
    carpeta_destino: str = "General/PoC Plantillas SaEGA",
) -> bool:
    """
    Sube un archivo a SharePoint mediante Microsoft Graph.
    - bytes_io: contenido del archivo (BytesIO).
    - nombre_archivo: nombre final (p.ej. '123456 - SaeGA.xlsm').
    - access_token: token Bearer válido.
    - hostname: tenant SharePoint.
    - site_name: path del sitio ('/sites/{site_name}').
    - carpeta_destino: ruta relativa de la carpeta dentro del Drive del sitio.

    Devuelve True si 200/201; en caso contrario imprime el error y devuelve False.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream",
    }

    # 1) Obtener siteId
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
    site_resp = requests.get(site_url, headers=headers)
    try:
        site_resp.raise_for_status()
    except Exception:
        print("Error obteniendo siteId:", site_resp.text)
        return False

    site_id = site_resp.json().get("id")
    if not site_id:
        print("No se pudo resolver el siteId. Respuesta:", site_resp.text)
        return False

    # 2) Subir archivo
    ruta_archivo = f"{carpeta_destino}/{nombre_archivo}"
    # Mantener las barras en la ruta (muy importante para Graph)
    ruta_archivo_enc = quote(ruta_archivo, safe="/")

    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_archivo_enc}:/content"
    data_bytes = bytes_io.getvalue() if hasattr(bytes_io, "getvalue") else bytes_io
    upload_resp = requests.put(upload_url, headers=headers, data=data_bytes)

    if upload_resp.status_code in (200, 201):
        return True

    print("Error al subir:", upload_resp.status_code, upload_resp.text)
    return False
