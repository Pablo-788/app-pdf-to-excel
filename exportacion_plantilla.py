# exportacion_plantilla.py
from io import BytesIO
import requests
from urllib.parse import quote
import os

try:
    # openpyxl >= 3.1 (para desplazar fórmulas al copiar)
    from openpyxl.formula.translate import Translator
except Exception:
    Translator = None


# ---- Ajusta esto si quieres un path por defecto para tu .xlsm local ----
RUTA_PLANTILLA_POR_DEFECTO = "SaeGA v2.0.2 - Plantilla - copia para Importador de Pedidos - copia.xlsm"

# COM de Excel
import pythoncom
import win32com.client as win32


def _as_2d(col_values):
    return [[v] for v in col_values]  # para asignación masiva COM


def limpiar_entradas_com(
    ruta_excel: str,
    hoja: str = "Pedidos",
    nombre_tabla: str = "tblPedidos",
    fila_inicio: int = 3,      # A3/B3/C3
    col_tienda: str = "A",
    col_referencia: str = "B",
    col_unidades: str = "C",
    ajustar_filas: bool = False  # si True, borra filas sobrantes de la tabla (deja 1)
) -> None:
    """
    Limpia SOLO las columnas de entrada (A/B/C) desde fila_inicio hacia abajo.
    Opcional: reduce filas de la tabla a 1 (sin tocar estilos/fórmulas).
    """
    pythoncom.CoInitialize()
    xl = None
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        # Evita que salten macros/eventos (error 424 por macros)
        try: xl.EnableEvents = False
        except: pass
        try: xl.AskToUpdateLinks = False
        except: pass
        try: xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except: pass

        wb = xl.Workbooks.Open(os.path.abspath(ruta_excel))
        ws = wb.Worksheets(hoja)

        last_row = fila_inicio  # por defecto

        # Si existe la tabla, usamos su tamaño para limpiar
        tbl = None
        try:
            tbl = ws.ListObjects(nombre_tabla)
            try:
                existing = int(tbl.DataBodyRange.Rows.Count)
            except Exception:
                existing = 0
            if existing > 0:
                first_data_row = tbl.DataBodyRange.Row
                last_row = first_data_row + existing - 1
            else:
                last_row = fila_inicio - 1
        except Exception:
            # Sin tabla: calculamos último usado en columna C (xlUp=-4162)
            last_row = ws.Cells(ws.Rows.Count, col_unidades).End(-4162).Row

        if last_row >= fila_inicio:
            ws.Range(f"{col_tienda}{fila_inicio}:{col_tienda}{last_row}").Value = ""
            ws.Range(f"{col_referencia}{fila_inicio}:{col_referencia}{last_row}").Value = ""
            ws.Range(f"{col_unidades}{fila_inicio}:{col_unidades}{last_row}").Value = ""

        # Opcional: reducir filas de la tabla a 1 (para dejarla limpia)
        if ajustar_filas and tbl is not None:
            try:
                existing = int(tbl.DataBodyRange.Rows.Count)
            except Exception:
                existing = 0
            # deja 1 fila de datos como plantilla
            for i in range(existing, 1, -1):
                tbl.ListRows(i).Delete()

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        if xl is not None:
            xl.Quit()
        pythoncom.CoUninitialize()


def exportar_directo_excel_com(
    ruta_excel: str,
    bytes_data: bytes,
    hoja: str = "Pedidos",
    nombre_tabla: str = "tblPedidos",
    fila_inicio: int = 3,      # A3/B3/C3
    col_tienda: str = "A",
    col_referencia: str = "B",
    col_unidades: str = "C",
    columnas_df = ("Tienda", "Código", "Cantidad"),
    modo: str = "sobrescribir"
) -> None:
    import os, pythoncom
    import pandas as pd
    import win32com.client as win32

    df = pd.read_excel(BytesIO(bytes_data)).copy()
    for col in columnas_df:
        if col not in df.columns:
            raise ValueError(f"Falta la columna '{col}' en el DataFrame de entrada.")
    n = len(df)
    if n == 0:
        return

    vals_tienda     = [[v] for v in df[columnas_df[0]].tolist()]
    vals_referencia = [[v] for v in df[columnas_df[1]].tolist()]
    vals_unidades   = [[v] for v in df[columnas_df[2]].tolist()]

    first_row = fila_inicio
    last_row  = fila_inicio + n - 1

    pythoncom.CoInitialize()
    xl = None
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try: xl.EnableEvents = False
        except: pass
        try: xl.AskToUpdateLinks = False
        except: pass
        try: xl.AutomationSecurity = 3
        except: pass

        wb = xl.Workbooks.Open(os.path.abspath(ruta_excel))
        ws = wb.Worksheets(hoja)

        # ===== AJUSTE CRÍTICO: que la tabla tenga EXACTAMENTE n filas =====
        tbl = None
        try:
            tbl = ws.ListObjects(nombre_tabla)
            try:
                existing = int(tbl.DataBodyRange.Rows.Count)
            except Exception:
                existing = 0

            if existing > n:
                # Borra filas sobrantes empezando desde abajo
                for i in range(existing, n, -1):
                    tbl.ListRows(i).Delete()
            elif existing < n:
                for _ in range(n - existing):
                    tbl.ListRows.Add()
        except Exception:
            tbl = None  # si no hay tabla, seguimos por rango

        # ===== ESCRITURA POR CELDAS A3/B3/C3 =====
        ws.Range(f"{col_tienda}{first_row}:{col_tienda}{last_row}").Value = vals_tienda
        ws.Range(f"{col_referencia}{first_row}:{col_referencia}{last_row}").Value = vals_referencia
        ws.Range(f"{col_unidades}{first_row}:{col_unidades}{last_row}").Value = vals_unidades

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        if xl is not None:
            xl.Quit()
        pythoncom.CoUninitialize()


def exportar_plantilla(
    bytes_data: bytes,
    ruta_excel: str = RUTA_PLANTILLA_POR_DEFECTO,
    **kwargs
) -> BytesIO:
    """
    Wrapper compatible con tu UI anterior:
      - Escribe en el .xlsm local (ruta_excel) usando exportar_directo_excel(...)
      - Devuelve el .xlsm actualizado como BytesIO para descargar/subir.
    """
    exportar_directo_excel_com(ruta_excel, bytes_data, **kwargs)
    return leer_plantilla_actualizada(ruta_excel)


def leer_plantilla_actualizada(ruta_excel: str) -> BytesIO:
    """Lee el .xlsm ya actualizado y devuelve un BytesIO listo para descargar/subir."""
    with open(ruta_excel, "rb") as f:
        data = f.read()
    bio = BytesIO(data)
    bio.seek(0)
    return bio


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
