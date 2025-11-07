import os
import subprocess
import time

# 📂 Rutas fijas (ajustá si cambiás ubicación)
carpeta = r"C:\Users\leandro.fonrtanar\OneDrive - AXEL JOHNSON INTERNATIONAL AB\Documents\VSCode\name_replacer\testing rig"
pdfcreator_path = r"C:\Program Files\PDFCreator\PDFCreator.exe"

sufijo = input("Ingrese el sufijo: ").strip()

# ⚙️ Reglas de renombrado
reglas = [
    ("ASK", "BOM", f"{sufijo} - cost (ASK)"),
    ("ASK", "Customer", f"{sufijo} - datasheet"),
    ("ASK", "INST", f"{sufijo} - Dimensions"),
    ("ASK", "PLATE", f"{sufijo} - Plate Arrangement"),
    ("01_", "Pricing", f"{sufijo} - cost (local)"),
]

# Validar PDFCreator
if not os.path.exists(pdfcreator_path):
    print("❌ No se encontró PDFCreator. Verifique la ruta.")
    exit()

# 🧾 1. Renombrar todos los archivos según las reglas
for archivo in os.listdir(carpeta):
    ruta_original = os.path.join(carpeta, archivo)
    if not os.path.isfile(ruta_original):
        continue

    nombre, ext = os.path.splitext(archivo)

    for prefijo, contiene, nuevo_nombre in reglas:
        if nombre.startswith(prefijo) and contiene in nombre:
            nuevo_archivo = nuevo_nombre + ext
            ruta_nueva = os.path.join(carpeta, nuevo_archivo)

            # Evitar sobrescribir si ya existe
            contador = 1
            while os.path.exists(ruta_nueva):
                base, extension = os.path.splitext(nuevo_archivo)
                ruta_nueva = os.path.join(carpeta, f"{base} ({contador}){extension}")
                contador += 1

            os.rename(ruta_original, ruta_nueva)
            print(f"Renombrado: {archivo} → {os.path.basename(ruta_nueva)}")
            break

# 🧾 2. Convertir solo los RTF relevantes a PDF
for archivo in os.listdir(carpeta):
    ruta = os.path.join(carpeta, archivo)
    if not (os.path.isfile(ruta) and archivo.lower().endswith(".rtf")):
        continue

    nombre, _ = os.path.splitext(archivo)

    # Detectar si corresponde a datasheet o Plate Arrangement
    if nombre.endswith("datasheet"):
        salida_pdf = os.path.join(carpeta, f"{sufijo} - datasheet.pdf")
    elif nombre.endswith("Plate Arrangement"):
        salida_pdf = os.path.join(carpeta, f"{sufijo} - Plate Arrangement.pdf")
    else:
        continue  # otros RTF no se convierten

    print(f"Convirtiendo {archivo} → {os.path.basename(salida_pdf)}")

    cmd = [
        pdfcreator_path,
        f"/PrintFile={ruta}",
        f"/OutputFile={salida_pdf}",
        "/Profile=DefaultGuid",
        "/NoStart",
        "/Close"
    ]
    subprocess.run(cmd)

    time.sleep(3)
    if os.path.exists(salida_pdf):
        print(f"✅ Creado: {os.path.basename(salida_pdf)}")
    else:
        print(f"⚠️ No se generó el PDF para {archivo}")

print("🏁 Proceso completado.")
