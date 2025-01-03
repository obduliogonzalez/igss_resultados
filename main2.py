
import pdfplumber
from openpyxl import Workbook

# Ruta del archivo PDF
archivo_pdf = "C:/Users/nestor.gonzalez/Documents/GitHub/igss_resultados/GARCIA_GAYTAN_OLGA_JUDITH_M.pdf"
#"C:\Users\nestor.gonzalez\Documents\GitHub\igss_resultados\GARCIA_GAYTAN_OLGA_JUDITH_M.pdf"

# Nombre del archivo Excel de salida
archivo_excel = "resultados_paciente.xlsx"

# Función para extraer texto de todas las páginas y estructurarlo
def extraer_datos_pdf(ruta_pdf):
    resultados = []
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                resultados.append(texto.strip())
    except Exception as e:
        print(f"Error al procesar el archivo: {e}")
    return resultados

# Función para procesar el texto y convertirlo en filas de Excel
def procesar_datos_para_excel(textos):
    datos = []
    for texto in textos:
        lineas = texto.split("\n")
        for linea in lineas:
            if ":" in linea:  # Filtrar las líneas con datos estructurados
                partes = linea.split(":")
                if len(partes) == 2:
                    clave, valor = partes[0].strip(), partes[1].strip()
                    datos.append([clave, valor])
                elif len(partes) > 2:  # Por si hay más de un ":" en la línea
                    clave = partes[0].strip()
                    valor = ":".join(partes[1:]).strip()
                    datos.append([clave, valor])
    return datos

# Extraer los datos del PDF
textos_paginas = extraer_datos_pdf(archivo_pdf)

# Procesar los textos para Excel
datos_excel = procesar_datos_para_excel(textos_paginas)

# Crear el archivo Excel
wb = Workbook()
ws = wb.active
ws.title = "Resultados"
ws.append(["Clave", "Valor"])  # Agregar encabezados

# Agregar los datos al Excel
for fila in datos_excel:
    ws.append(fila)

# Guardar el archivo Excel
wb.save(archivo_excel)
print(f"Archivo Excel guardado como: {archivo_excel}")
