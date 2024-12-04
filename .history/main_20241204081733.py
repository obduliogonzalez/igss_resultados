
import os
import pdfplumber
import re
from openpyxl import Workbook

# Ruta de la carpeta con los archivos PDF
carpeta_pdf = "C:/Users/nestor.gonzalez/Documents/GitHub/igss_resultados"

import os
import pdfplumber
import re
from openpyxl import Workbook

# Ruta de la carpeta con los archivos PDF
carpeta_pdf = "C:/Users/nestor.gonzalez/Documents/GitHub/destino_resultados"

# Ruta de salida para los archivos renombrados (puede ser la misma carpeta)
ruta_salida = "ruta/a/tu/carpeta/con/pdfs"

# Nombre del archivo Excel de salida
archivo_excel = "resultados_pacientes.xlsx"

# Expresión regular para encontrar nombres de pacientes
patron_nombre = re.compile(r"Paciente:\s*([A-Z\s]+)")

# Función para extraer el nombre del paciente usando expresión regular
def extraer_nombre_paciente(ruta_pdf):
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            # Leer la primera página
            pagina = pdf.pages[0]
            texto = pagina.extract_text()

            # Buscar el nombre del paciente usando la expresión regular
            coincidencia = patron_nombre.search(texto)
            if coincidencia:
                return coincidencia.group(1).strip()
            else:
                return "Nombre no encontrado"
    except Exception as e:
        return f"Error al leer el archivo: {e}"

# Crear un libro de Excel
wb = Workbook()
ws = wb.active
ws.title = "Resultados"
ws.append(["Archivo", "Nombre del Paciente"])  # Agregar encabezados

# Iterar sobre todos los archivos PDF en la carpeta
for archivo in os.listdir(carpeta_pdf):
    if archivo.endswith(".pdf"):
        ruta_archivo = os.path.join(carpeta_pdf, archivo)
        nombre_paciente = extraer_nombre_paciente(ruta_archivo)

        # Agregar los resultados al archivo de Excel de forma ordenada
        ws.append([archivo, nombre_paciente])

        # Renombrar el archivo PDF si se extrajo un nombre válido
        if nombre_paciente != "Nombre no encontrado":
            # Reemplazar caracteres no permitidos en nombres de archivos
            nombre_paciente = nombre_paciente.replace(" ", "_").replace("/", "-").replace("\\", "-")
            nuevo_nombre = f"{nombre_paciente}.pdf"
            ruta_nueva = os.path.join(ruta_salida, nuevo_nombre)

            # Evitar sobrescribir archivos si ya existe un archivo con el nuevo nombre
            contador = 1
            while os.path.exists(ruta_nueva):
                nuevo_nombre = f"{nombre_paciente}_{contador}.pdf"
                ruta_nueva = os.path.join(ruta_salida, nuevo_nombre)
                contador += 1

            # Renombrar el archivo
            os.rename(ruta_archivo, ruta_nueva)
            print(f"Archivo renombrado: {archivo} -> {nuevo_nombre}")

# Guardar el archivo Excel
wb.save(archivo_excel)
print(f"Resultados guardados en {archivo_excel}")
