from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, NamedStyle, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
import pandas as pd
import time
import pyautogui
import os
import glob
import random
import xlwings as xw
import pdfkit
import fitz

# Definir rutas a las carpetas y archivos
input_folder_excel = "C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/input/Deudas"
output_folder_csv = "C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/input/DeudasCSV"
output_file_csv = "C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/Resumen_deudas.csv"
output_file_xlsx = "C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/Resumen_deudas.xlsx"
fecha_especifica = '2024-03-31'

# Leer el archivo Excel
df = pd.read_excel(r'C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/input/clientes.xlsx')

# Suposición de nombres de columnas
cuit_login_list = df['CUIT para ingresar'].tolist()
cuit_represent_list = df['CUIT representado'].tolist()
password_list = df['Contraseña'].tolist()
download_list = df['Ubicacion Descarga'].tolist()
posterior_list = df['Posterior'].tolist()
anterior_list = df['Anterior'].tolist()
clientes_list = df['Cliente'].tolist()

output_folder_pdf = "C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/Reportes"
imagen = "C:/Program Files/Sublime Merge/Descarga-SCT-Envio-Mails-Masivos/data/imagen.png"

def forzar_guardado_excel(excel_file):
    app = xw.App(visible=False)
    wb = app.books.open(excel_file)
    wb.save()
    wb.close()
    app.quit()

def procesar_excel(excel_file, output_pdf, imagen):
    # Cargar el archivo Excel con pandas
    df = pd.read_excel(excel_file)

    # Filtrar por "Periodo fiscal" y "Impuesto"
    df_filtrado = df[
        (df['Período Fiscal'].astype(str).str.contains('2024')) & 
        (df['Impuesto'].str.contains('ganancias personas físicas|ganancias personas fisicas|bienes personales', case=False, na=False))
    ]

    # Verificar si la tabla está vacía
    if df_filtrado.shape[0] == 0:
        output_pdf = output_pdf.replace(".pdf", " - vacio.pdf")

    # Eliminar las columnas innecesarias
    if 'Concepto / Subconcepto' in df.columns:
        df_filtrado = df_filtrado.drop(['Concepto / Subconcepto'], axis=1)

    if 'Int. resarcitorios' in df.columns:
        df_filtrado = df_filtrado.drop(['Int. resarcitorios'], axis=1)

    if 'Int. punitorios' in df.columns:
        df_filtrado = df_filtrado.drop(['Int. punitorios'], axis=1)

    # Guardar el DataFrame filtrado en el archivo Excel
    df_filtrado.to_excel(excel_file, index=False)
    
    # Cargar el archivo para aplicar formato con openpyxl
    wb = load_workbook(excel_file)
    ws = wb.active

    # Insertar filas adicionales si necesitas espacio adicional para una nueva imagen
    ws.insert_rows(1, amount=7)

    # Agregar una imagen encima del encabezado (A1)
    img = ExcelImage(imagen)
    total_width = sum([ws.column_dimensions[get_column_letter(col)].width for col in range(1, ws.max_column + 1)])
    img.width = total_width * 7
    img.height = 120
    ws.add_image(img, 'A1')

    # Cambiar el color del encabezado a lila
    header_fill = PatternFill(start_color="AA0EAA", end_color="AA0EAA", fill_type="solid")
    for cell in ws[8]:
        cell.fill = header_fill

    # Ajustar el ancho de las columnas automáticamente
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width

    # Centrar el contenido de todas las celdas
    for row in ws.iter_rows(min_row=8, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # Guardar los cambios
    wb.save(excel_file)

    # Iniciar una instancia de xlwings de forma invisible
    app = xw.App(visible=False)
    try:
        wb_xw = app.books.open(excel_file)
        last_row = ws.max_row
        last_col = ws.max_column
        wb_xw.sheets[0].api.PageSetup.PrintArea = f"A1:{get_column_letter(last_col)}{last_row}"

        wb_xw.sheets[0].api.PageSetup.Zoom = False
        wb_xw.sheets[0].to_pdf(output_pdf)
        wb_xw.close()
    finally:
        app.quit()

    print(f"Archivo {excel_file} procesado y guardado como {output_pdf}")


# Recorrer todos los archivos Excel en la carpeta
for excel_file in glob.glob(os.path.join(input_folder_excel, "*.xlsx")):
    try:
        # Forzar guardado para evitar problemas con archivos corruptos o no calculados
        forzar_guardado_excel(excel_file)

        # Obtener el nombre base del archivo para usarlo en el nombre del PDF
        base_name = os.path.splitext(os.path.basename(excel_file))[0]
        output_pdf = os.path.join(output_folder_pdf, f"{base_name}.pdf")
        
        # Llamar a la función para procesar el archivo Excel y generar el PDF
        procesar_excel(excel_file, output_pdf, imagen)
        
        print(f"Archivo {excel_file} procesado y guardado como {output_pdf}")
    
    except Exception as e:
        print(f"Error al procesar {excel_file}: {e}")