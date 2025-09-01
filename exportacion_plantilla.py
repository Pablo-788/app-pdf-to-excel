import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
import requests
from urllib.parse import quote

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

    estilos_columna = {
        idx: {
            "font": copy(celda.font),
            "border": copy(celda.border),
            "fill": copy(celda.fill),
            "number_format": copy(celda.number_format),
            "alignment": copy(celda.alignment),
        }
        for idx, celda in enumerate(fila_ejemplo)
    }

    # 6️⃣ Rellenar datos y copiar estilos
    for i, fila_origen in enumerate(df_origen.itertuples(index=False), start=0):
        for col_df, col_tabla in mapeo_columnas.items():
            if col_df in df_origen.columns and col_tabla in encabezados_tabla:
                idx_col = encabezados_tabla.index(col_tabla)
                celda_destino = ws.cell(row=fila_ejemplo[0].row + i, column=idx_col + 1)

                # Rellenar valor
                celda_destino.value = getattr(fila_origen, col_df)

                # Copiar estilos de la fila de ejemplo
                for attr, val in estilos_columna[idx_col].items():
                    setattr(celda_destino, attr, copy(val))

    # 7️⃣ Guardar a BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def subir_a_sharepoint(bytes_io: BytesIO, nombre_archivo: str, access_token: str) -> bool:
# 📌 Configuración
    hostname = "saboraespana.sharepoint.com"
    site_name = "departamento.ti"
    carpeta_destino = "General/PoC Plantillas SaEGA"  # Ruta relativa dentro del sitio

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream"
    }

    # 1️⃣ Obtener el siteId
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    # 2️⃣ Subir el archivo al destino
    # Construimos la ruta completa del archivo
    ruta_archivo = f"{carpeta_destino}/{nombre_archivo}"
    ruta_archivo = quote(ruta_archivo)

    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_archivo}:/content"
    upload_resp = requests.put(upload_url, headers=headers, data=bytes_io.getvalue())

    # 3️⃣ Verificar resultado
    if upload_resp.status_code in [200, 201]:
        return True
    else:
        print("Error al subir:", upload_resp.text)
        return False