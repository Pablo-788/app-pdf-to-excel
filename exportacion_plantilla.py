# exportacion_plantilla.py
from io import BytesIO
import requests
from urllib.parse import quote
import os
import pandas as pd
import json
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ---- Ajusta esto si quieres un path por defecto para tu .xlsm local ----
# Ya no es necesario, pero lo mantenemos como recordatorio
RUTA_PLANTILLA_POR_DEFECTO = "SaeGA v2.0.2 - Plantilla - copia para Importador de Pedidos - copia.xlsm"

def obtener_drive_id_y_item_id(
    access_token: str,
    hostname: str = "saboraespana.sharepoint.com",
    site_name: str = "departamento.ti",
    carpeta_destino: str = "General/PoC Plantillas SaEGA",
) -> tuple[str, str] | None:
    """
    Obtiene el driveId y itemId de la carpeta destino en SharePoint.
    - access_token: token Bearer válido.
    - hostname: tenant SharePoint.
    - site_name: path del sitio ('/sites/{site_name}').
    - carpeta_destino: ruta relativa de la carpeta dentro del Drive del sitio.

    Devuelve una tupla (driveId, itemId) o None si hay error.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    # 1) Obtener siteId
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
    site_resp = requests.get(site_url, headers=headers)
    try:
        site_resp.raise_for_status()
    except Exception:
        print("Error obteniendo siteId:", site_resp.text)
        return None
    site_id = site_resp.json().get("id")
    if not site_id:
        print("No se pudo resolver el siteId. Respuesta:", site_resp.text)
        return None
    
    # 2) Obtener driveId
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drive_resp = requests.get(drive_url, headers=headers)
    try:
        drive_resp.raise_for_status()
    except Exception:
        print("Error obteniendo driveId:", drive_resp.text)
        return None
    drives = drive_resp.json().get("value", [])
    
    # Suponiendo que hay un único drive por defecto, o que buscamos uno específico
    # Esta parte de tu código está incompleta, así que la ajustamos para obtener el drive por defecto
    if not drives:
        print("No se encontraron drives.")
        return None
    drive_id = drives[0].get('id') # Toma el primer drive, que suele ser el de "Documentos"
    
    # 3) Obtener itemId de la carpeta destino
    # Aquí debemos asegurarnos de que la ruta de la carpeta destino es correcta
    # Para el ámbito de este ejemplo, asumiremos que la plantilla ya está en el drive, 
    # y la buscaremos por nombre. El código de tu pregunta no implementa esta parte,
    # así que la omitimos para centrarnos en el problema principal.
    # El itemId debe ser proporcionado de alguna manera. En un escenario real,
    # podrías buscarlo por su nombre dentro de la carpeta.
    print("La función 'obtener_drive_id_y_item_id' no está completa en el código original.")
    print("Deberías implementar la lógica para encontrar el 'itemId' de tu archivo Excel.")
    return None # Devolvemos None ya que la lógica no está implementada aquí
import requests

def limpiar_entradas_mgraph(access_token: str, drive_id: str, item_id: str, nombre_tabla: str) -> None:
    """
    Borra todas las filas de una tabla de Excel excepto la primera fila de datos.

    Args:
        drive_id (str): El ID del Drive donde se encuentra el archivo.
        item_id (str): El ID del Item (archivo Excel) dentro del Drive.
        nombre_tabla (str): El nombre de la tabla dentro del libro de Excel.
        access_token (str): El token de acceso para autenticarse en Microsoft Graph API.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # 1. CONSTRUIR LA URL PARA OBTENER LAS FILAS Y CONTARLAS
    url_filas = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
        f"workbook/tables('{nombre_tabla}')/rows"
    )

    try:
        # 2. OBTENER LA LONGITUD DE LA TABLA
        response = requests.get(url_filas, headers=headers)
        response.raise_for_status()  # Lanza un error si la petición falla (e.g., 401, 404)
        
        filas = response.json().get('value', [])
        longitud_tabla = len(filas)
        logging.info(f"La tabla '{nombre_tabla}' tiene {longitud_tabla} fila(s) de datos.")

        # 3. VERIFICAR SI HAY ALGO QUE BORRAR
        if longitud_tabla <= 1:
            logging.info("La tabla tiene una fila o ninguna. No se requiere ninguna acción.")
            return

        # 4. BORRAR LAS FILAS EN UN BUCLE
        num_filas_a_borrar = longitud_tabla - 1
        logging.info(f"Se procederá a borrar {num_filas_a_borrar} fila(s)...")

        # La URL para borrar es siempre la misma: la fila en el índice 1
        url_borrado = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
            f"workbook/tables('{nombre_tabla}')/rows/itemAt(index=1)"
        )

        for i in range(num_filas_a_borrar):
            delete_response = requests.delete(url_borrado, headers=headers)
            
            # El código 204 "No Content" indica un borrado exitoso
            if delete_response.status_code == 204:
                logging.info(f"Fila {i + 1}/{num_filas_a_borrar} borrada con éxito.")
            else:
                logging.error(f"Error al borrar la fila {i + 1}. "
                              f"Status: {delete_response.status_code}, "
                              f"Respuesta: {delete_response.text}")
                # Si una falla, es mejor detenerse para investigar
                break
        
        logging.info("Proceso de limpieza de la tabla finalizado.")

    except requests.exceptions.HTTPError as http_err:
        logging.error(f"Error HTTP: {http_err} - {response.text}")
    except Exception as err:
        logging.error(f"Ocurrió un error inesperado: {err}")



  
def col_to_excel_letter(col_num: int) -> str:
    """Converts a 1-based column number to its Excel letter representation."""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string

def escribir_datos_mgraph(
    access_token: str,
    drive_id: str,
    item_id: str,
    bytes_data: bytes,
    hoja: str = "Pedidos",
    celda_inicio: str = "A3",
    columnas_df: tuple = ("Tienda", "Código", "Cantidad")
) -> None:
    """
    Escribe datos de un DataFrame en una hoja de Excel en SharePoint 
    usando la API de Microsoft Graph.
    """
    df = pd.read_excel(BytesIO(bytes_data))
    if df.empty:
        print("DataFrame vacío, no hay datos para escribir.")
        return
        
    try:
        df_to_write = df[list(columnas_df)]
    except KeyError as e:
        raise ValueError(f"Falta una columna requerida en el DataFrame: {e}")

    # Convertir el DataFrame a una lista de listas para el payload de Graph
    values = df_to_write.values.tolist()
    
    # === CORRECTED LOGIC ===
    # Get the start row and column from celda_inicio
    # This assumes a simple 'A3' or 'B5' format
    start_col_letter = "".join(filter(str.isalpha, celda_inicio))
    start_row = int("".join(filter(str.isdigit, celda_inicio)))
    
    # Get the 1-based column number for the start letter
    start_col_num = 0
    for i, char in enumerate(start_col_letter):
        start_col_num = start_col_num * 26 + (ord(char.upper()) - ord('A')) + 1

    # Calculate the number of rows and columns to determine the end cell.
    num_rows = len(df_to_write)
    num_cols = len(df_to_write.columns)

    # Calculate the end row and column
    end_row = start_row + num_rows - 1
    end_col_num = start_col_num + num_cols - 1
    
    # Convert the end column number to its Excel letter representation
    end_col_letter = col_to_excel_letter(end_col_num)

    # Construct the full range address
    full_range = f"{celda_inicio}:{end_col_letter}{end_row}"

    # Construir la URL del endpoint 'update range'
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
        f"workbook/worksheets('{hoja}')/range(address='{full_range}')"
    )
    # ========================
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    # El payload debe tener la estructura {"values": [[...]]}
    payload = {
        "values": values
    }
    
    print(f"Enviando petición PATCH para escribir datos a la URL: {url}")
    try:
        response = requests.patch(url, headers=headers, json=payload)
        response.raise_for_status()
        print("¡Los datos han sido escritos exitosamente en el Excel!")
       
    except requests.exceptions.HTTPError as err:
        print(f"Error HTTP al escribir datos: {err}")
        print(f"Respuesta del servidor: {response.text}")
        raise
    except Exception as err:
        print(f"Ocurrió un error inesperado al escribir datos: {err}")
        raise


import requests

def limpiar_filas_vacias_despues_datos(
    access_token: str,
    drive_id: str,
    item_id: str,
    hoja: str = "Pedidos"
) -> None:
    """
    Detecta la primera tabla de una hoja y elimina cualquier fila que esté
    completamente vacía, dondequiera que se encuentre.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json',
    }
    
    try:
        # 1. Detectar automáticamente la primera tabla en la hoja
        print(f"⚙️ Buscando la tabla en la hoja '{hoja}'...")
        url_list_tables = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
            f"workbook/worksheets('{hoja}')/tables"
        )
        response_tables = requests.get(url_list_tables, headers=headers)
        response_tables.raise_for_status()
        
        tables = response_tables.json().get('value', [])
        if not tables:
            print(f"❌ No se encontraron tablas en la hoja '{hoja}'.")
            return
            
        nombre_tabla = tables[0]['name']
        print(f"👍 Tabla '{nombre_tabla}' encontrada.")

        # 2. Obtener todas las filas de la tabla
        url_get_rows = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
            f"workbook/tables('{nombre_tabla}')/rows"
        )
        response_rows = requests.get(url_get_rows, headers=headers)
        response_rows.raise_for_status()
        filas_tabla = response_rows.json().get('value', [])
        
        if not filas_tabla:
            print("ℹ️ La tabla no tiene filas de datos.")
            return

        # 3. Identificar los índices de todas las filas que están completamente vacías
        indices_a_eliminar = []
        for i, fila in enumerate(filas_tabla):
            valores = fila.get('values', [[]])[0]
            # Si todas las celdas de la fila están vacías, la marcamos para eliminar
            if not any(str(celda).strip() for celda in valores):
                indices_a_eliminar.append(i)

        if not indices_a_eliminar:
            print("✅ No se encontraron filas vacías para eliminar.")
            return
            
        # 4. Decidir la acción: LIMPIAR o ELIMINAR
        # Si el número de filas a eliminar es igual al total de filas, no podemos
        # usar DELETE. En su lugar, debemos LIMPIAR el contenido de la tabla.
        if len(indices_a_eliminar) == len(filas_tabla):
            print("⚠️ Todas las filas están vacías. Se limpiará el contenido de la tabla.")
            url_clear_body = (
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
                f"workbook/tables('{nombre_tabla}')/dataBodyRange/clear"
            )
            payload = {"applyTo": "All"}
            response_clear = requests.post(url_clear_body, headers=headers, json=payload)
            response_clear.raise_for_status()
            print("✨ Contenido de la tabla limpiado exitosamente.")
        else:
            # Si solo algunas filas están vacías, las eliminamos.
            # Se itera en orden inverso para no alterar los índices de las filas pendientes.
            print(f"🗑️ Se eliminarán {len(indices_a_eliminar)} filas vacías...")
            # Elimina las filas vacías en orden inverso para evitar problemas de índice
            for idx in reversed(indices_a_eliminar):
                url_delete_row = (
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/"
                    f"workbook/tables('{nombre_tabla}')/rows/itemAt(index=1)"
                )
                response_delete = requests.delete(url_delete_row, headers=headers)
                try:
                    response_delete.raise_for_status()
                    print(f"✅ Fila vacía en índice {idx} eliminada correctamente.")
                except requests.exceptions.HTTPError as err:
                    print(f"❌ Error al eliminar la fila en índice {idx}: {err}")
                    print(f"Respuesta del servidor: {response_delete.text}")
                    raise

    except requests.exceptions.HTTPError as err:
        print(f"❌ Error HTTP: {err}")
        if hasattr(err, 'response') and err.response:
            print(f"    Respuesta del servidor: {err.response.text}")
        raise
    except Exception as err:
        print(f"❌ Error inesperado: {err}")
        raise

def copy_file_on_sharepoint(
    access_token: str,
    drive_id: str,
    item_id: str,
    new_file_name: str
) -> dict | None:
    """
    Copies a file within the same SharePoint drive and folder.
    
    Args:
        access_token (str): A valid Microsoft Graph access token.
        drive_id (str): The ID of the SharePoint drive (e.g., the Documents library).
        item_id (str): The ID of the file to be copied.
        new_file_name (str): The name for the new copy (e.g., 'Copia_factura.xlsx').

    Returns:
        A dictionary with the response from the API or None if an error occurs.
    """
    copy_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/copy"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    payload = {
        'name': new_file_name
    }
    
    try:
        response = requests.post(copy_url, headers=headers, json=payload)
        response.raise_for_status()  # Raise an exception for HTTP errors (4xx or 5xx)
        
        # The copy operation is asynchronous.
        # The API returns a 'Location' header to check the status.
        return response.json()
        
    except requests.exceptions.HTTPError as err:
        print(f"Que ese nombre de fichero no esté ya en el servidor! - Error HTTP: {err}")
        print(f"Server response: {response.text}")
        return None
    except Exception as err:
        print(f"An unexpected error occurred: {err}")
        return None

def descargar_file_de_sharepoint(
    access_token: str,
    drive_id: str,
    item_id: str,
) -> BytesIO | None:
    """
    Descarga el contenido de un archivo de SharePoint usando su item_id.
    """
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    try:
        response = requests.get(url, headers=headers, stream=True)
        response.raise_for_status()

        # Check the Content-Type to avoid trying to parse HTML as JSON
        if 'text/html' in response.headers.get('Content-Type', ''):
            print("Received HTML response instead of file content. Token might be invalid or expired.")
            print(f"Server response: {response.text[:500]}...") # Print a snippet of the HTML
            # You should implement token refresh logic here
            return None
        
        # Stream the content to handle potentially large files
        file_content = BytesIO()
        for chunk in response.iter_content(chunk_size=8192):
            file_content.write(chunk)
        file_content.seek(0)
        
        return file_content
        
    except requests.exceptions.HTTPError as err:
        print(f"Error HTTP: {err}")
        print(f"Server response: {response.text}")
        return None
    except Exception as err:
        print(f"An unexpected error occurred: {err}")
        return None
