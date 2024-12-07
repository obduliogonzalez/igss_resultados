

# Encabezados para la tabla
headers = [
    "ID", "Sexo", "RBC-ERITROCITOS", "HGB-HEMOGLOBINA", "HCT-HEMATOCRITO", "MCV", "MCH", "MCHC", 
    "WBC-LEUCOCITOS", "LYM%-LINFOCITOS", "NEU%-NEUTROFILOS", "MON%-MONOCITOS", 
    "EOS%-EOSINOFILOS", "BAS%-BASOFILOS", "LYN#-LINFOCITOS", "NEU#-NEUTROFILOS", 
    "MON#-MONOCITOS", "EOS#-EOSINOFILOS", "BAS#-BASOFILOS", "RDW-CV", "RDW-SD", 
    "PLT-PLAQUETAS", "MPV", "PDW"
]

# Escribir encabezados en Excel
for col_num, header in enumerate(headers, 1):
    sheet.cell(row=1, column=col_num, value=header)

# Función para buscar y extraer datos relevantes
def extract_data(text):
    """
    Busca las líneas clave en el texto y extrae los datos.
    """
    data = {}
    for line in text.split('\n'):
        parts = line.split()
        if not parts:  # Ignorar líneas vacías
            continue
        try:
            if "ID:" in line:
                data["ID"] = line.split(":")[-1].strip()
            elif "Sexo:" in line:
                data["Sexo"] = line.split(":")[-1].strip()
            elif "RBC-ERITROCITOS" in line:
                data["RBC-ERITROCITOS"] = parts[1]
            elif "HGB-HEMOGLOBINA" in line:
                data["HGB-HEMOGLOBINA"] = parts[1]
            elif "HCT-HEMATOCRITO" in line:
                data["HCT-HEMATOCRITO"] = parts[1]
            elif "MCV" in line and len(parts) > 1:
                data["MCV"] = parts[1]
            elif "MCH" in line and len(parts) > 1:
                data["MCH"] = parts[1]
            elif "MCHC" in line and len(parts) > 1:
                data["MCHC"] = parts[1]
            elif "WBC-LEUCOCITOS" in line:
                data["WBC-LEUCOCITOS"] = parts[1]
            elif "LYM%-LINFOCITOS" in line:
                data["LYM%-LINFOCITOS"] = parts[1]
            elif "NEU%-NEUTROFILOS" in line:
                data["NEU%-NEUTROFILOS"] = parts[1]
            elif "MON%-MONOCITOS" in line:
                data["MON%-MONOCITOS"] = parts[1]
            elif "EOS%-EOSINOFILOS" in line:
                data["EOS%-EOSINOFILOS"] = parts[1]
            elif "BAS%-BASOFILOS" in line:
                data["BAS%-BASOFILOS"] = parts[1]
            elif "LYN#-LINFOCITOS" in line:
                data["LYN#-LINFOCITOS"] = parts[1]
            elif "NEU#-NEUTROFILOS" in line:
                data["NEU#-NEUTROFILOS"] = parts[1]
            elif "MON#-MONOCITOS" in line:
                data["MON#-MONOCITOS"] = parts[1]
            elif "EOS#-EOSINOFILOS" in line:
                data["EOS#-EOSINOFILOS"] = parts[1]
            elif "BAS#-BASOFILOS" in line:
                data["BAS#-BASOFILOS"] = parts[1]
            elif "RDW-CV" in line:
                data["RDW-CV"] = parts[1]
            elif "RDW-SD" in line:
                data["RDW-SD"] = parts[1]
            elif "PLT-PLAQUETAS" in line:
                data["PLT-PLAQUETAS"] = parts[1]
            elif "MPV" in line:
                data["MPV"] = parts[1]
            elif "PDW" in line:
                data["PDW"] = parts[1]
        except IndexError:
            print(f"Error procesando línea: {line}")
            continue
    return data

# Iterar sobre cada página del PDF y extraer datos
row_num = 2
for page_num in range(len(pdf)):
    page = pdf.load_page(page_num)
    text = page.get_text("text")  # Extraer el texto
    data = extract_data(text)  # Extraer datos relevantes
    if data:
        for col_num, header in enumerate(headers, 1):
            sheet.cell(row=row_num, column=col_num, value=data.get(header, ""))
        row_num += 1

# Guardar el archivo Excel
output_file = 'resultados_procesados.xlsx'
workbook.save(output_file)
print(f"Archivo Excel generado: {output_file}")