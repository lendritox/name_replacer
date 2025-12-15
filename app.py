import os
import string
import shutil
import subprocess
import time
import zipfile
import re
import win32com.client as win32
import datetime


# 📂 CONFIGURACIÓN

# BASE_DIR = r"C:\Users\leandro.fonrtanar\OneDrive - AXEL JOHNSON INTERNATIONAL AB\Documents\VSCode\name_replacer" # for testing
BASE_DIR = r"C:\Users\leandro.fonrtanar\OneDrive\Documentos\20_Laburo\Brown Brothers\00_QUOTATIONS\04_DESIGNS"
PLANTILLA_QUOTE = r"C:\Users\leandro.fonrtanar\OneDrive\Documentos\20_Laburo\Brown Brothers\00_QUOTATIONS\00_Quotations Utilities\00_Quotation.xlsm"  # ajustá la ruta
PDFCREATOR_PATH = r"C:\Program Files\PDFCreator\PDFCreator.exe"

# ----------------------------------------------------------
# 🔧 FUNCIONES DE UTILIDAD
# ----------------------------------------------------------



def check_folder(numero):
    base_dir = BASE_DIR
    carpeta_existente = None
    for carpeta in os.listdir(base_dir):
        if carpeta.startswith(str(numero)):
            carpeta_existente = os.path.join(base_dir, carpeta)
            break

    if carpeta_existente:
        opcion = input("La carpeta ya existe desea procesar archivos?").strip()
        if opcion == "y":
            procesar_cotizacion(carpeta_existente)
        else:
            print("❌ Operación cancelada.")
            return  
    else:
        crear_estructura_cotizacion(numero)
    
def crear_estructura_cotizacion(numero):
    """Crea la carpeta madre y subcarpetas para una nueva cotización."""
    empresa = input("Empresa: ").strip()
    contacto = input("Contacto: ").strip()
    contacto_carpeta = contacto.split()[0]
    descripcion = input("Descripción general: ").strip()
    codigo_ask = input("Código ASK: ").strip()
    n_items = int(input("Cantidad de ítems a diseñar: ").strip())

    # Crear carpeta madre
    carpeta_madre = os.path.join(BASE_DIR, f"{numero} - {empresa} - {contacto_carpeta} - {descripcion} - {codigo_ask}")
    os.makedirs(carpeta_madre, exist_ok=True)
    print(f"📁 Creada carpeta principal: {carpeta_madre}")

    # Crear subcarpetas
    for i in range(n_items):
        letra = string.ascii_lowercase[i]
        sub = os.path.join(carpeta_madre, f"{numero}.{letra} - ")
        os.makedirs(sub, exist_ok=True)
        print(f"  ➕ Subcarpeta: {sub}")

    # Copiar plantilla de cotización y modifcarla
    if os.path.exists(PLANTILLA_QUOTE):
        destino = os.path.join(carpeta_madre, f"{numero}.01 - quote.xlsm")
        shutil.copy(PLANTILLA_QUOTE, destino)
        print(f"🧾 Copiada plantilla de cotización: {destino}")
        
        modificar_excel(destino, numero, empresa, contacto, n_items)
    
    else:
        print("⚠️ No se encontró la plantilla de cotización.")

    print("✅ Estructura de cotización creada.")

def modificar_excel(destino, numero, empresa, contacto, n_items):

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # ponelo en True para debug
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(destino)
    ws = wb.Worksheets(1)

    # === 1) Escribir datos generales ===
    ws.Range("C4").Value = empresa.title()
    ws.Range("C5").Value = contacto.title()

    # año en 2 dígitos
    year2 = str(datetime.datetime.now().year % 100).zfill(2)
    ws.Range("G3").Value = f"{year2}.{numero}.01"

    # === 2) Preparar la tabla de items ===
    # La tabla arranca en fila 10, los encabezados están en 9
    start_row = 10

    # Cantidad total de filas necesarias:
    total_new_rows = n_items * 3  # local + imported + shipment por cada item

    # Insertamos filas ANTES de la fila 10
    # Excel desplaza perfectamente estilos, alturas, bordes, imágenes, TODO.
    ws.Rows(f"{start_row + 1}:{start_row + total_new_rows + 1}").Insert()

    # === 3) Llenar filas ===
    current_row = start_row

    for i in range(1, n_items + 1):

        # Part number base
        pn_base = f"{numero}.{chr(96+i)}"  # a, b, c ...

        # Fila 1: Local
        ws.Range(f"A{current_row}").Value = current_row-9
        ws.Range(f"B{current_row}").Value = f"{pn_base} (local)"
        ws.Range(f"C{current_row}").Value = ""
        ws.Range(f"D{current_row}").Value = 1
        ws.Range(f"E{current_row}").Value = "5-6 weeks"
        ws.Range(f"F{current_row}").Value = ""  # precio unitario (completás luego)
        ws.Range(f"G{current_row}").Formula = f"=F{current_row}*D{current_row}"

        current_row += 1

        # Fila 2: Imported
        ws.Range(f"A{current_row}").Value = current_row-9
        ws.Range(f"B{current_row}").Value = f"{pn_base} (imported)"
        ws.Range(f"C{current_row}").Value = ""
        ws.Range(f"D{current_row}").Value = 0
        ws.Range(f"E{current_row}").Value = "TBC"
        ws.Range(f"F{current_row}").Value = ""
        ws.Range(f"G{current_row}").Formula = f"=F{current_row}*D{current_row}"

        current_row += 1

        # Fila 3: Shipment
        ws.Range(f"A{current_row}").Value = current_row-9
        ws.Range(f"B{current_row}").Value = "-"
        ws.Range(f"C{current_row}").Value = "DHL shipment from factory"
        ws.Range(f"D{current_row}").Value = 0
        ws.Range(f"E{current_row}").Value = "5-6 weeks"
        ws.Range(f"F{current_row}").Value = 2300
        ws.Range(f"G{current_row}").Formula = f"=F{current_row}*D{current_row}"

        # Agregar borde inferior grueso a esta fila
        border = ws.Range(f"A{current_row}:G{current_row}").Borders(9)  # xlEdgeBottom = 9
        border.Weight = 2  # xlMedium

        current_row += 1

    # === 4) Guardar ===
    wb.Save()
    wb.Close()
    excel.Quit()

    print("Excel modificado correctamente.")

def procesar_cotizacion(carpeta_madre):
    """Procesa todas las subcarpetas de una cotización existente."""

    print(f"📁 Procesando cotización: {carpeta_madre}")

    # Recorre subcarpetas
    for sub in os.listdir(carpeta_madre):
        sub_path = os.path.join(carpeta_madre, sub)
        
        # patrón: 15007.a o 15007.a - descripción
        if os.path.isdir(sub_path) and re.match(rf"^{re.escape(numero)}\.[^ \-]+(?:\s-\s.*)?$", sub, re.IGNORECASE):
            sufijo = sub.split(" - ")[0]  # Ej: "15007.a"
            print(f"\n➡️ Procesando {sufijo}")
            renombrar_y_convertir(sub_path, sufijo)

    print("\n✅ Cotización procesada para envío.")

def renombrar_y_convertir(carpeta, sufijo):
    """Renombra, convierte y empaqueta los archivos de una carpeta de ítem."""
    reglas = [
        ("ASK", "BOM", f"{sufijo} - cost (ASK)"),
        ("ASK", "Customer", f"{sufijo} - datasheet"),
        ("ASK", "INST", f"{sufijo} - Dimensions"),
        ("ASK", "PLATE", f"{sufijo} - Plate Arrangement"),
        ("01_", "Pricing", f"{sufijo} - cost (local)"),
    ]

    # 1. Renombrar
    for archivo in os.listdir(carpeta):
        ruta_original = os.path.join(carpeta, archivo)
        if not os.path.isfile(ruta_original):
            continue
        nombre, ext = os.path.splitext(archivo)
        for prefijo, contiene, nuevo_nombre in reglas:
            if nombre.startswith(prefijo) and contiene in nombre:
                nuevo_archivo = nuevo_nombre + ext
                ruta_nueva = os.path.join(carpeta, nuevo_archivo)
                contador = 1
                while os.path.exists(ruta_nueva):
                    base, extension = os.path.splitext(nuevo_archivo)
                    ruta_nueva = os.path.join(carpeta, f"{base} ({contador}){extension}")
                    contador += 1
                os.rename(ruta_original, ruta_nueva)
                print(f"  🔤 Renombrado: {archivo} → {os.path.basename(ruta_nueva)}")
                break

    # 2. Convertir RTF → PDF
    for archivo in os.listdir(carpeta):
        ruta = os.path.join(carpeta, archivo)
        if not (os.path.isfile(ruta) and archivo.lower().endswith(".rtf")):
            continue
        nombre, _ = os.path.splitext(archivo)
        if nombre.endswith("datasheet"):
            salida_pdf = os.path.join(carpeta, f"{sufijo} - datasheet.pdf")
        elif nombre.endswith("Plate Arrangement"):
            salida_pdf = os.path.join(carpeta, f"{sufijo} - Plate Arrangement.pdf")
        else:
            continue

        print(f"  🖨️ Convirtiendo {archivo} → {os.path.basename(salida_pdf)}")
        cmd = [
            PDFCREATOR_PATH,
            f"/PrintFile={ruta}",
            f"/OutputFile={salida_pdf}",
            "/Profile=DefaultGuid",
            "/NoStart",
            "/Close"
        ]

        # eliminar PDF previo si ya existe
        if os.path.exists(salida_pdf):
            os.remove(salida_pdf)
            print(f"  ♻️ Eliminado PDF anterior: {os.path.basename(salida_pdf)}")

        subprocess.run(cmd)
        time.sleep(3)

    # 3. Crear ZIP si existen los dos archivos clave
    pdf = os.path.join(carpeta, f"{sufijo} - datasheet.pdf")
    dxf = os.path.join(carpeta, f"{sufijo} - Dimensions.dxf")
    if os.path.exists(pdf):
        carpeta_superior = os.path.dirname(carpeta)
        zip_destino = os.path.join(carpeta_superior, f"{sufijo}.zip")

        # eliminar ZIP previo si ya existe
        if os.path.exists(zip_destino):
            os.remove(zip_destino)
            print(f"  ♻️ Eliminado ZIP anterior: {os.path.basename(zip_destino)}")

        if os.path.exists(dxf):
            with zipfile.ZipFile(zip_destino, 'w', compression=zipfile.ZIP_STORED) as zipf:
                zipf.write(pdf, os.path.basename(pdf))
                zipf.write(dxf, os.path.basename(dxf))
        else:
            print(f"⚠️ No DXG {sufijo}")    
            with zipfile.ZipFile(zip_destino, 'w', compression=zipfile.ZIP_STORED) as zipf:
                zipf.write(pdf, os.path.basename(pdf))
        print(f"  📦 ZIP creado: {zip_destino}")
    else:
        print("  ⚠️ NO DATASHEET")


# ----------------------------------------------------------
# 🚀 MENÚ PRINCIPAL
# ----------------------------------------------------------

if __name__ == "__main__":
    print("=== Gestor de Cotizaciones PHE ===")
    
    numero = input("Número de cotización: ").strip()
    check_folder(numero)
    

