
import pdfplumber  # Asegúrate de importar este módulo
# Ruta al archivo PDF
archivo_pdf = "C:/Users/nestor.gonzalez/Documents/GitHub/igss_resultados"

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

# Llamar a la función y mostrar el nombre
nombre = extraer_nombre_paciente(archivo_pdf)
print(f"Nombre del paciente: {nombre}")
