import os
import pdfplumber

# Ruta donde se encuentran tus archivos PDF
ruta_pdf = "C:/ruta/a/tus/pdf"  # Cambia esto a la carpeta que contiene los archivos


# Ruta de destino para los archivos renombrados
ruta_destino = "C:/ruta/a/destino"  # Cambia esto a la carpeta de destino

# Crear la carpeta de destino si no existe
os.makedirs(ruta_destino, exist_ok=True)

# Función para manejar conflictos de nombres
def generar_nombre_unico(carpeta, nombre_base, extension):
    contador = 1
    nuevo_nombre = f"{nombre_base}{extension}"
    while os.path.exists(os.path.join(carpeta, nuevo_nombre)):
        nuevo_nombre = f"{nombre_base}_{contador}{extension}"
        contador += 1
    return nuevo_nombre

# Procesar los archivos en la carpeta
for archivo in os.listdir(ruta_pdf):
    if archivo.endswith(".pdf"):
        ruta_archivo = os.path.join(ruta_pdf, archivo)
        try:
            # Abrir el archivo PDF y leer el contenido
            with pdfplumber.open(ruta_archivo) as pdf:
                # Leer el texto de la primera página
                pagina = pdf.pages[0]
                texto = pagina.extract_text()

                # Buscar el nombre del paciente
                if "Paciente:" in texto:
                    inicio = texto.find("Paciente:") + len("Paciente:")
                    fin = texto.find("Muestra:", inicio)
                    nombre_paciente = texto[inicio:fin].strip()

                    # Crear el nuevo nombre del archivo
                    nombre_base = nombre_paciente.replace(" ", "_")  # Reemplazar espacios por guiones bajos
                    nuevo_nombre = generar_nombre_unico(ruta_destino, nombre_base, ".pdf")
                    ruta_nueva = os.path.join(ruta_destino, nuevo_nombre)

                    # Renombrar el archivo
                    os.rename(ruta_archivo, ruta_nueva)
                    print(f"Renombrado: {archivo} -> {nuevo_nombre}")
                else:
                    print(f"No se encontró 'Paciente:' en el archivo {archivo}")
        except Exception as e:
            print(f"Error al procesar {archivo}: {e}")
