import os

# Pedir carpeta y sufijo
carpeta = input("Ingrese la ruta de la carpeta: ").strip()
sufijo = input("Ingrese el sufijo: ").strip()

# Validar carpeta
if not os.path.isdir(carpeta):
    print("❌ La carpeta especificada no existe.")
    exit()

# Definir patrones y nombres nuevos
reglas = [
    ("ASK", "BOM", f"{sufijo} - cost (ASK)"),
    ("ASK", "Customer", f"{sufijo} - Datasheet"),
    ("ASK", "INST", f"{sufijo} - Dimensions"),
    ("ASK", "PLATE", f"{sufijo} - Plate Arrangement"),
    ("01_", "Pricing", f"{sufijo} - cost (local)"),
]

# Procesar archivos
for archivo in os.listdir(carpeta):
    ruta_original = os.path.join(carpeta, archivo)
    if not os.path.isfile(ruta_original):
        continue  # ignorar carpetas

    nombre, ext = os.path.splitext(archivo)

    for prefijo, contiene, nuevo_nombre in reglas:
        if nombre.startswith(prefijo) and contiene in nombre:
            nuevo_archivo = nuevo_nombre + ext
            ruta_nueva = os.path.join(carpeta, nuevo_archivo)

            # Evitar sobrescribir archivos
            contador = 1
            while os.path.exists(ruta_nueva):
                base, extension = os.path.splitext(nuevo_archivo)
                ruta_nueva = os.path.join(carpeta, f"{base} ({contador}){extension}")
                contador += 1

            os.rename(ruta_original, ruta_nueva)
            print(f"Renombrado: {archivo} → {os.path.basename(ruta_nueva)}")
            break

print("✅ Proceso completado.")
