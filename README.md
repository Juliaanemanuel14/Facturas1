import os
import json
import pandas as pd
import gradio as gr
import tempfile
import zipfile
import time
import pythoncom
import win32com.client
import pdfplumber
import fitz  # PyMuPDF
from google import genai
from dotenv import load_dotenv
from PIL import Image
import re
from datetime import datetime 

# Cargar variables de entorno
load_dotenv()

api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    raise ValueError("❌ ERROR: No se encontró la clave API en el archivo .env")

client = genai.Client(api_key=api_key)

BATCH_SIZE = 10
MAX_RETRIES = 3

def limpiar_nombre_proveedor(file_name):
    """Limpia el nombre del archivo para extraer el nombre del proveedor."""
    nombre = os.path.splitext(file_name)[0]  # Eliminar la extensión
    nombre = re.sub(r'_pagina_\d+', '', nombre)  # Eliminar 'pagina_x'
    nombre = re.sub(r'[_\d]+', ' ', nombre)  # Eliminar números y guiones bajos
    nombre = nombre.strip().title()  # Convertir a formato de título
    return nombre

def extraer_archivos(archivo_zip):
    temp_dir = tempfile.mkdtemp()
    pdf_files, excel_files, image_files = [], [], []
    
    try:
        with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
    except zipfile.BadZipFile:
        print("❌ ERROR: El archivo ZIP está corrupto o no es válido.")
        return [], [], [], temp_dir

    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
            elif file.endswith(('.xls', '.xlsx')):
                excel_files.append(os.path.join(root, file))
            elif file.endswith(('.png', '.jpg', '.jpeg')):
                image_files.append(os.path.join(root, file))

    return pdf_files, excel_files, image_files, temp_dir

def excel_a_pdf(ruta_excel, carpeta_salida):
    """Convierte un archivo Excel en PDF."""
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(ruta_excel)

        if workbook.ActiveSheet:
            hoja = workbook.ActiveSheet
        else:
            hoja = workbook.Sheets(workbook.Sheets.Count)

        hoja.PageSetup.Zoom = False
        hoja.PageSetup.FitToPagesWide = 1
        hoja.PageSetup.FitToPagesTall = False

        ruta_pdf = os.path.join(carpeta_salida, f"{os.path.basename(ruta_excel).replace('.xlsx', '.pdf')}")
        hoja.ExportAsFixedFormat(0, ruta_pdf)

        workbook.Close(False)
        excel.Quit()
        
        print(f"✅ PDF generado desde Excel: {ruta_pdf}")
        return ruta_pdf
    except Exception as e:
        print(f"❌ ERROR al convertir {ruta_excel} a PDF: {e}")
        return None

def dividir_pdf_por_paginas(pdf_path):
    """Divide un PDF en archivos individuales por cada página solo si tiene más de una página."""
    paginas = []
    doc = fitz.open(pdf_path)
    if len(doc) == 1:
        return [pdf_path]
    
    for i in range(len(doc)):
        temp_path = f"{pdf_path[:-4]}_pagina_{i+1}.pdf"
        new_pdf = fitz.open()
        new_pdf.insert_pdf(doc, from_page=i, to_page=i)
        new_pdf.save(temp_path)
        paginas.append(temp_path)
    return paginas

def procesar_zip(archivo_zip):
    pdf_files, excel_files, image_files, temp_dir = extraer_archivos(archivo_zip)
    all_data = []

    # Convertir Excel a PDF antes de procesar
    for excel_file in excel_files:
        print(f"🔄 Convirtiendo Excel a PDF: {excel_file}")
        pdf_generado = excel_a_pdf(excel_file, temp_dir)
        if pdf_generado:
            print(f"✅ PDF generado: {pdf_generado}")
            pdf_files.append(pdf_generado)
        else:
            print(f"❌ Fallo al convertir: {excel_file}")

    # Asegurarse de procesar todos los archivos PDF, incluidos los generados desde Excel
    archivos_a_procesar = pdf_files + image_files
    for file_path in archivos_a_procesar:
        paginas_pdf = dividir_pdf_por_paginas(file_path) if file_path.endswith(".pdf") else [file_path]
        
        for pagina in paginas_pdf:
            file_name = os.path.basename(pagina)
            proveedor = limpiar_nombre_proveedor(file_name)
            print(f"📄 Procesando: {file_name} (Proveedor: {proveedor})")
            
            try:
                uploaded_file = client.files.upload(file=pagina, config={'display_name': file_name})
                print(f"🚀 Archivo enviado a Gemini: {file_name}")
            except Exception as e:
                print(f"❌ ERROR al subir {file_name} a Gemini: {e}")
                continue
            
            prompt = f"""
Extrae la información de los archivos PDF indicados y devuélvela en formato JSON con los siguientes campos:

- **Fecha**: formato (d-mmm-yy)
- **Fecha de entrega**: formato (d-mmm-yy), Que sea la fecha de recepcion de la mercaderia
- **Codigo Proveedor**: Puede que algunos archivos no lo tengan, dejar vacio en ese caso
- **Cantidad**: Valor numerico
- **Articulo o Producto**: Nombre completo del artículo, especificando presentación o unidad si corresponde.
- **Precio**: Valor del precio con separador de miles (Ejemplo: 1.000 o 5.000.000).
- **% descuento**: Valor %
- **Descuento 1**: Valor del descuento con separador de miles (Ejemplo: 1.000 o 5.000.000).
- **Descuento 2**: Valor del descuento con separador de miles (Ejemplo: 1.000 o 5.000.000).
- **Precio sin impuestos**: Valor del precio con separador de miles (Ejemplo: 1.000 o 5.000.000).
- **Total por producto**: Valor del total con separador de miles (Ejemplo: 1.000 o 5.000.000).
- **Proveedor**: tiene que ser el de la factura. Asigna el nombre '{proveedor}' como proveedor.


---

### **Reglas Generales:**
- Corrige errores de formato en los precios (separadores incorrectos, símbolos no válidos, etc.).
- Elimina caracteres especiales que puedan causar errores:
  - Reemplaza comillas dobles `"` por comillas simples `'`.
  - Elimina o corrige cualquier símbolo no válido (como `\`, saltos de línea o caracteres no imprimibles).
- Los saltos de línea dentro de los campos deben ser reemplazados por espacios.
- Extrae todos los artículos sin omitir ninguno, incluso si el precio está incompleto.

---

### **Reglas Específicas por Archivo:**


- **Asegúrate de extraer todos los artículos presentes en el archivo, sin omitir ninguno.**

### **Validación Final:**
- Verifica que todos los artículos del archivo hayan sido extraídos.
- Asegúrate de que los precios estén correctamente estructurados y sin errores de formato..
- El JSON final debe ser válido, bien estructurado y sin errores antes de entregarlo.
"""
            
            try:
                response = client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=[prompt, uploaded_file],
                    config={"response_mime_type": "application/json"}
                )
                raw_response = response.candidates[0].content.parts[0].text
                json_data = json.loads(raw_response)
                all_data.extend(json_data)
            except Exception as e:
                print(f"❌ ERROR en la solicitud a Gemini para {file_name}: {e}")
                continue
    
    df = pd.DataFrame(all_data)
    df["Fecha"] = datetime.today().strftime("%d/%m/%Y")
    excel_path = os.path.join(temp_dir, "datos_extraidos.xlsx")
    df.to_excel(excel_path, index=False)
    
    return excel_path

with gr.Blocks() as app:
    gr.Markdown("# 📄 Procesador de Archivos desde ZIP con Gemini")
    gr.Markdown("Sube un archivo ZIP con PDFs, Excels o Imágenes y obtendrás un Excel con los datos extraídos.")

    with gr.Row():
        input_file = gr.File(label="Subir archivo ZIP", file_types=[".zip"])
        output_file = gr.File(label="Descargar Excel", interactive=True)
    
    submit_button = gr.Button("Procesar") # Botón para procesar el archivo
    submit_button.click(procesar_zip, inputs=input_file, outputs=output_file)

app.launch(server_port=4218, share=True)
