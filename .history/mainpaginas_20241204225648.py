import fitz  # PyMuPDF

# Ruta al archivo PDF original
input_pdf_path = 'C:/Users/nestor.gonzalez/Documents/GitHub/igss_resultados/GARCIA_GAYTAN_OLGA_JUDITH_1.pdf'
output_pdf_path = '/mnt/data/archivo_filtrado.pdf'

# Crear una lista con las páginas que queremos imprimir (1, 4, 7, ..., 1603)
pages_to_print = [i for i in range(1, 1604) if (i - 1) % 3 == 0]

# Abrir el archivo PDF original
input_pdf = fitz.open(input_pdf_path)

# Crear un nuevo documento PDF
output_pdf = fitz.open()

# Agregar solo las páginas seleccionadas al nuevo PDF
for page_num in pages_to_print:
    if page_num <= len(input_pdf):  # Verificar que la página exista en el PDF original
        page = input_pdf.load_page(page_num - 1)  # El índice es cero-based
        output_pdf.insert_pdf(input_pdf, from_page=page_num - 1, to_page=page_num - 1)

# Guardar el archivo PDF filtrado
output_pdf.save(output_pdf_path)

print(f"El archivo PDF filtrado se ha guardado en: {output_pdf_path}")
