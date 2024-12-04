
import os
import pdfplumber  # Asegúrate de importar este módulo
# Ruta al archivo PDF
carpeta_pdf = "C:/Users/nestor.gonzalez/Documents/GitHub/igss_resultados"


# Función para extraer el nombre del paciente
def extraer_nombre_paciente(ruta_pdf):
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            # Leer la primera página
            pagina = pdf.pages[0]
            texto = pagina.extract_text()

            # Buscar el nombre del paciente
            if "Paciente:" in texto:
                inicio = texto.find("Paciente:") + len("Paciente:")
                fin = texto.find("Muestra:", inicio)
                nombre_paciente = texto[inicio:fin].strip()
                return nombre_paciente
            else:
                return "Nombre no encontrado"
    except Exception as e:
        return f"Error al leer el archivo: {e}"

# Iterar sobre todos los archivos PDF en la carpeta
for archivo in os.listdir(carpeta_pdf):
    if archivo.endswith(".pdf"):
        ruta_archivo = os.path.join(carpeta_pdf, archivo)
        nombre_paciente = extraer_nombre_paciente(ruta_archivo)
        print(f"Archivo: {archivo} -> Nombre del paciente: {nombre_paciente}")
        
        