import os
import re
import xlsxwriter
import pdfplumber  # Use pdfplumber for better table handling

# Function to scan a PDF for keywords, including content in tables
def scan_pdf_with_pdfplumber(file_path, keywords):
    results = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                # Extract plain text and table content
                page_text = page.extract_text()
                table_content = page.extract_tables()

                # Join all table rows into one large block of text for scanning
                table_text = ""
                for table in table_content:
                    for row in table:
                        row_text = " ".join([str(cell) for cell in row if cell])  # Combine row cells into one string
                        table_text += row_text + " "

                # Combine both regular text and table text
                full_text = (page_text or "") + "\n" + table_text.strip()

                if not full_text.strip():  # Skip empty pages
                    continue

                # Split text into sentences using regex
                sentences = re.split(r'(?<=[.!?])\s+', full_text)

                # Count occurrences of each keyword (including multi-word ones) using regex
                for keyword in keywords:
                    # Escape special characters in the keyword
                    keyword_pattern = r'\b' + re.escape(keyword) + r'\b'

                    for sentence in sentences:
                        # Check if the keyword exists in the sentence (case-insensitive search)
                        keyword_count = len(re.findall(keyword_pattern, sentence, re.IGNORECASE))
                        if keyword_count > 0:
                            results.append({
                                'file': os.path.basename(file_path),
                                'keyword': keyword,
                                'count': keyword_count,
                                'sentence': sentence.strip(),
                                'page': page_number  # Include the page number
                            })

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
    return results

# Directory containing PDF files
pdf_directory = "/Users/sophiawinter/Desktop/All Spanish NDCs"

# Define keyword categories

keyword_categories = {

    "Nature-Based Solutions": [
        "manglares", "manglar", "praderas de pastos marinos", "pradera de pastos marinos",
        "arrecifes de coral", "arrecife de coral", "bosques de algas pardas", "bosque de algas pardas",
        "bancos de ostras", "banco de ostras", "macroalgas", "lodos intermareales", "lodo intermareal",
        "marismas salinas", "marisma salina", "marismas de marea", "marisma de marea",
        "restauración de manglares", "ecosistemas de manglares"
    ],

    "Ocean-Based Mitigation": [
        "mitigación basada en el océano", "mitigación basada en el océano",
        "energía mareomotriz", "energía undimotriz", "energía eólica marina",
        "energía eólica marina", "calor oceánico", "carbono azul",
        "CCS basado en el océano", "captura y almacenamiento de carbono basado en el océano",
        "CCS azul", "captura y almacenamiento de carbono azul", "captura de carbono azul",
        "almacenamiento de carbono azul", "solar flotante",
        "descarbonización del transporte marítimo", "secuestro de carbono marino",
        "secuestro de carbono oceánico", "ecosistemas de carbono azul",
        "energías renovables oceánicas", "energía renovable oceánica", "energía oceánica",
        "tecnología de energía marina", "tecnologías de energía marina",
        "energía renovable en alta mar", "tecnologías marítimas verdes",
        "tecnologías oceánicas", "tecnología basada en el océano",
        "tecnologías basadas en el océano", "área marina", "áreas marinas",
        "eliminación de dióxido de carbono marino", "transporte marítimo de cero emisiones",
        "CDR oceánico", "eliminación de dióxido de carbono oceánico",
        "descarbonización del transporte marítimo", "transporte marítimo verde",
        "tecnología de pesca", "fuentes de carbono azul", "restauración de carbono azul",
        "descarbonización de los sectores basados en el océano"
    ],

    "Ocean-Based Adaptation": [
        "adaptación basada en el océano", "adaptación basada en el océano",
        "diversificación oceánica", "diversificación marina", "diversificación azul",
        "sistemas de alerta temprana marina", "sistema de alerta temprana marina",
        "áreas marinas protegidas", "área marina protegida", "protección de aguas profundas",
        "AMPs", "AMP", "gestión integrada de zonas costeras",
        "ICZM", "ICZMs", "sistemas de alerta temprana oceánicos",
        "sistema de alerta temprana oceánico", "EWS marino", "EWS oceánico",
        "ABMT", "herramientas de gestión basadas en áreas",
        "evaluaciones de riesgo clima-océano", "adaptación al aumento del nivel del mar",
        "contabilidad oceánica", "cuentas oceánicas", "aumento del nivel del mar",
        "finanzas oceánicas", "30 by 30", "30 x 30", "minería en aguas profundas",
        "década del océano", "ciencia oceánica", "conferencia de la ONU sobre el océano",
        "Marco Global de Biodiversidad de Kunming-Montreal", "bonos azules", "finanzas azules",
        "contaminación plástica marina", "recursos genéticos marinos",
        "áreas más allá de la jurisdicción nacional", "acuerdo BBNJ",
        "acuerdo sobre biodiversidad más allá de la jurisdicción nacional"
    ],

    "Sustainable Fisheries": [
        "pesquerías sostenibles", "pesquería sostenible", "producción de alimentos azules",
        "alimentos azules", "seguridad alimentaria acuática", "producción de alimentos acuáticos",
        "alimentos acuáticos", "pesca a pequeña escala", "peces", "pesquerías",
        "producción de alimentos sostenible", "derechos de pesca", "industria marítima",
        "industrias marítimas", "acuicultura"
    ],

    "Sustainable Ocean Plans": [
        "planes oceánicos sostenibles", "plan oceánico sostenible", "SOPs",
        "planificación oceánica sostenible basada en la ciencia", "gestión de vías fluviales",
        "plan oceánico", "planes oceánicos", "investigación oceánica"
    ],

    "Marine Spatial Planning": [
        "planificación espacial marina", "MSPs", "plan espacial marino", "planes espaciales marinos"
    ],

    "Illegal, Unreported and Unregulated Fishing": [
        "IUU", "pesca ilegal, no declarada y no reglamentada", "pesca IUU",
        "buques pesqueros ilegales", "buque pesquero ilegal"
    ],

    "General ocean-related Mentions": [
        "océano", "inundaciones", "inundación", "inundación", "marino", "costa",
        "costero", "protección costera", "acción climática basada en el océano",
        "nivel del agua costera", "azul", "marítimo", "mar", "gobernanza costera",
        "entornos marinos", "entorno marino", "biodiversidad marina",
        "ciencia oceánica", "gestión oceánica", "resiliencia costera",
        "el océano", "ecosistemas oceánicos", "ecosistema oceánico",
        "alfabetización oceánica", "enfoque basado en ecosistemas",
        "comunidades costeras", "PSSA", "Áreas Marinas particularmente sensibles",
        "restauración oceánica", "servicio de acuicultura marina",
        "servicios de acuicultura marina", "flotas pesqueras",
        "protección del océano", "sistema de pastos marinos",
        "sistemas de pastos marinos", "hábitats de la línea de costa",
        "hábitat de la línea de costa", "hábitats submarinos",
        "hábitat submarino", "aumento del nivel del mar",
        "restauración de la biodiversidad oceánica",
        "fortalecimiento de la resiliencia costera",
        "acidificación oceánica", "gestión de desechos oceánicos",
        "bienestar de los animales acuáticos", "humedales costeros",
        "humedal costero", "protección marina",
        "inventarios nacionales de humedales", "inventario nacional de humedales",
        "ecosistemas costeros", "ecosistema costero", "acidificación costera",
        "acidificación", "desoxigenación oceánica", "águas profundas",
        "alta mar", "planificación marina inteligente frente al clima",
        "nuevas fuentes de carbono azul", "geoingeniería oceánica",
        "de la fuente al mar", "seguridad alimentaria marina",
        "evaluaciones de riesgo clima-océano", "animal acuático",
        "investigación científica marina",
        "oportunidades de empleo resilientes marinas y costeras",
        "planificación oceánica sostenible basada en la ciencia",
        "restauración y protección de carbono azul",
        "informe especial del IPCC sobre cambio climático y criosfera",
        "criosfera", "capa polar", "datos de pronóstico oceánico",
        "hielo marino ártico", "ecosistemas costeros en inventarios de GEI",
        "observación oceánica", "tecnologías de observación oceánica",
        "océano y derechos humanos", "conocimiento oceánico",
        "datos oceánicos", "brechas de datos oceánicos", "estresores marinos",
        "geoingeniería marina", "soluciones basadas en el océano",
        "cambios oceánicos inducidos por el ser humano y el clima",
        "eventos meteorológicos extremos"
    ],

    "Ocean Observation": [
        "inventarios de humedales", "inventario de humedales", "datos oceánicos",
        "inventario nacional de humedales", "ciencia marina", "cartografía de ecosistemas",
        "contabilidad nacional", "contabilidad de carbono azul",
        "metodologías de contabilidad de carbono azul", "mercado de carbono azul",
        "observación sistemática basada en el océano", "observación basada en el océano",
        "observación oceánica", "datos de pronóstico oceánico", "observación marítima"
    ],

    "SDG 14": [
        "ODS 14", "contaminación marina", "proteger los ecosistemas costeros",
        "proteger los ecosistemas marinos", "reducir la acidificación oceánica",
        "pesca sostenible", "conservación de áreas costeras",
        "conservación de áreas marinas", "conservar lo costero", "conservar lo marino",
        "subsidios a la pesca", "subsidios que contribuyen a la sobrepesca",
        "uso sostenible de los recursos marinos", "pesquerías de pequeña escala",
        "aumentar el conocimiento para la salud del océano", "salud del océano",
        "derecho marítimo internacional", "ambición azul",
        "Plan de Adaptación al Cambio Climático Costero", "CCCAP",
        "Estrategias y Planes de Acción Nacional de Biodiversidad", "EPANB", "EPANB"
    ],

    "Ocean Finance": [
        "finanzas oceánicas", "inversión oceánica", "inversión relacionada con el océano",
        "inversiones relacionadas con el océano", "gestión de recursos marinos",
        "financiación de carbono azul", "financiación de la restauración de ecosistemas"
    ],

    "Sustainable Ocean-Based Tourism": [
        "turismo oceánico sostenible", "turismo costero sostenible", "turismo costero",
        "turismo oceánico", "contaminación marina", "bonos azules", "infraestructura azul",
        "sector económico oceánico"
    ],

    "Ocean-related Storms": [
        "depresión tropical", "depresiones tropicales",
        "tormenta tropical", "tormentas tropicales",
        "huracán", "huracanes",
        "tifón", "tifones",
        "ciclón", "ciclones",
        "ciclón subtropical", "ciclones subtropicales",
        "ciclón extratropical", "ciclones extratropicales",
        "tormenta nor'easter", "tormentas nor'easter",
        "tormenta eólica europea", "tormentas eólicas europeas",
        "baja polar", "bajas polares",
        "medicán", "medicanes",
        "tormenta kona", "tormentas kona",
        "sistema convectivo de mesoescala", "sistemas convectivos de mesoescala",
        "línea de turbonada", "líneas de turbonada",
        "tromba marina tornádica", "trombas marinas tornádicas",
        "tromba marina de buen tiempo", "trombas marinas de buen tiempo"
    ],

}


# Scan the directory
def scan_pdfs_in_directory(directory, keyword_categories):
    all_results = []
    for file_name in os.listdir(directory):
        if file_name.lower().endswith('.pdf'):
            file_path = os.path.join(directory, file_name)
            print(f"Processing file: {file_name}")
            for category, keywords in keyword_categories.items():
                try:
                    # Scan the file for keywords
                    results = scan_pdf_with_pdfplumber(file_path, keywords)
                    all_results.extend(results)
                except Exception as e:
                    print(f"Error scanning file {file_name} for category {category}: {e}")
    return all_results

# Results
results = scan_pdfs_in_directory(pdf_directory, keyword_categories)

# Remove duplicate entries
unique_results = []
seen = set()
for r in results:
    # Create a unique key from the key fields
    key = (r['file'], r['page'], r['keyword'], r['count'], r['sentence'])
    if key not in seen:
        seen.add(key)
        unique_results.append(r)

results = unique_results

# Save results to Excel using xlsxwriter
output_excel_path = "/Users/sophiawinter/Desktop/spanischfloods.xlsx"
workbook = xlsxwriter.Workbook(output_excel_path)
worksheet = workbook.add_worksheet("Keyword Results")

# Add headers
headers = ["File", "Page", "Keyword", "Count", "Sentence"]
header_format = workbook.add_format({'bold': True})
for col_num, header in enumerate(headers):
    worksheet.write(0, col_num, header, header_format)

# Write results and apply bold to actual keywords in sentences
row_num = 1
for result in results:
    # Add basic data
    worksheet.write(row_num, 0, result['file'])
    worksheet.write(row_num, 1, result['page'])
    worksheet.write(row_num, 2, result['keyword'])
    worksheet.write(row_num, 3, result['count'])

    # Apply bold formatting to keywords within sentences
    sentence = result['sentence']
    keyword = result['keyword']
    keyword_pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)

    # Create segments for rich text formatting
    segments = []
    last_end = 0

    for match in keyword_pattern.finditer(sentence):
        start, end = match.start(), match.end()
        if start > last_end:  # Add text before the keyword
            segments.append(sentence[last_end:start])
        # Add the keyword with bold formatting
        segments.append(workbook.add_format({'bold': True}))
        segments.append(sentence[start:end])
        last_end = end

    if last_end < len(sentence):  # Add any remaining text after the last match
        segments.append(sentence[last_end:])

    # Write the rich text with formatting
    worksheet.write_rich_string(row_num, 4, *segments)
    row_num += 1

# Close workbook
workbook.close()

print(f"Data successfully saved to {output_excel_path}")