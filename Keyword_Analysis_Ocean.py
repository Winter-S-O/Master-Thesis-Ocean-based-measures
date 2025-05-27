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
        "mangroves", "mangrove", "seagrass beds", "seagrass bed", "coral reefs", "coral reef",
        "kelp forests", "kelp forest", "oyster beds", "oyster bed", "macroalgae", "mudflats", "mudflat",
        "salt marshes", "salt marsh", "tidal marshes", "tidal marsh", "mangrove restoration", "mangrove ecosystems"
    ],

    "Ocean-Based Mitigation": [
        "ocean-based mitigation", "ocean based mitigation", "tidal power", "wave power",
        "offshore wind", "offshore wind energy", "ocean heat", "blue carbon",
        "ocean based CCS", "ocean based carbon capture and storage", "blue CCS",
        "blue carbon capture and storage", "blue carbon capture", "blue carbon storage",
        "floating solar", "decarbonization of ocean based transport", "marine carbon sequestration",
        "ocean carbon sequestration", "blue carbon ecosystems", "renewable ocean energies",
        "renewable ocean energy", "ocean energy", "marine energy technology",
        "marine energy technologies", "offshore renewable energy", "green maritime technologies",
        "ocean technologies", "ocean-based technology", "ocean-based technologies", "marine area",
        "marine areas", "marine carbon dioxide removal", "zero-emission shipping", "ocean CDR",
        "ocean carbon dioxide removal", "decarbonization of shipping", "green shipping",
        "fishing technology", "blue carbon sources", "blue carbon restoration",
        "decarbonization of ocean-based sectors"
    ],

    "Ocean-Based Adaptation": [
        "ocean-based adaptation", "ocean based adaptation", "ocean diversification",
        "marine diversification", "blue diversification", "marine early warning systems",
        "marine early warning system", "marine protected areas", "marine protected area",
        "deep-sea protection", "MPAs", "MPA", "integrated coastal zone management", "ICZM", "ICZMs",
        "ocean early warning systems", "ocean early warning system", "marine EWS",
        "ocean EWS", "ABMT", "Area-based Management Tools", "climate-ocean risk assessments",
        "adaptation to sea level rise", "ocean accounting", "ocean accounts", "sea level rise",
        "ocean finance", "30 by 30", "30 x 30", "deep sea mining", "ocean decade", "ocean science",
        "UN ocean conference", "Kunming-Montreal Global Biodiversity Framework",
        "blue bonds", "blue finance", "marine plastic pollution", "marine genetic resources",
        "Areas beyond national jurisdiction", "BBNJ agreement",
        "Biodiversity Beyond National Jurisdiction agreement"
    ],

    "Sustainable Fisheries": [
        "sustainable fisheries", "sustainable fishery", "blue food production", "blue food",
        "aquatic food security", "aquatic food production", "aquatic food", "small-scale fishing",
        "fish", "fisheries", "sustainable food production", "fishing rights", "maritime industry",
        "maritime industries", "aquaculture"
    ],

    "Sustainable Ocean Plans": [
        "sustainable ocean plans", "sustainable ocean plan", "SOPs",
        "science-based sustainable ocean planning", "waterways management", "ocean plan",
        "ocean plans", "ocean research"
    ],

    "Marine Spatial Planning": [
        "marine spatial planning", "MSPs", "marine spatial plan", "marine spatial plans"
    ],

    "Illegal, Unreported and Unregulated Fishing": [
        "IUU", "Illegal, unreported and unregulated fishing", "IUU fishing",
        "illegal fishing vessels", "illegal fishing vessel"
    ],
    "General ocean-related Mentions": [
        "ocean", "floods", "flood", "flooding",
        "ocean-based climate action", "marine", "coast", "coastal"
        "coastal water level", "blue", "maritime", "sea", "coastal governance",
        "marine environments", "marine environment", "marine biodiversity", "ocean science",
        "ocean management", "coastal resilience", "the ocean", "ocean ecosystems",
        "ocean ecosystem", "ocean literacy", "ecosystem-based approach", "coastal communities",
        "PSSA", "Particularly Sensitive Sea Areas", "ocean restoration",
        "marine aquaculture service", "marine aquaculture services", "fish fleets",
        "ocean protection", "seagrass system", "seagrass systems", "shoreline habitats",
        "shoreline habitat", "underwater habitats", "underwater habitat", "sea level rise",
        "ocean biodiversity restoration", "strengthening coastal resilience", "ocean acidification",
        "ocean waste management", "aquatic animal welfare", "coastal wetlands", "coastal wetland",
        "marine protection", "national wetlands inventories", "national wetlands inventory",
        "coastal ecosystems", "coastal ecosystem", "coastal acidification", "acidification",
        "ocean deoxygenation", "deep sea", "high seas", "climate smart marine planning",
        "new blue carbon sources", "ocean geoengineering", "source-to-sea", "marine food security",
        "climate-ocean risk assessments", "aquatic animal", "marine scientific research",
        "Marine and coastal resilient employment opportunities",
        "science-based sustainable ocean planning", "blue carbon restoration and protection",
        "IPCC special report on climate change and cryosphere",
        "cryosphere", "polar sheet", "ocean forecast data", "arctic sea ice",
        "coastal ecosystems in GHG inventories", "ocean observation",
        "ocean observation technologies", "ocean and human rights", "ocean knowledge",
        "ocean data", "ocean data gaps", "marine stressors", "marine geoengineering",
        "ocean-based solutions", "human- and climate-induced ocean changes",
        "extreme weather events",
    ],

    "Ocean Observation": [
        "wetland inventories", "wetland inventory", "ocean data", "national wetland inventory",
        "marine science", "ecosystem mapping", "national accounting", "blue carbon accounting",
        "blue carbon accounting methodologies", "blue carbon market",
        "ocean-based systematic observation", "ocean-based observation", "ocean observation",
        "ocean forecast data", "maritime observation"
    ],

    "SDG 14": [
        "SDG 14", "marine pollution", "protect coastal ecosystems", "protect marine ecosystems",
        "reduce ocean acidification", "sustainable fishing", "conservation of coastal areas",
        "conservation of marine areas", "conserve coastal", "conserve marine",
        "fisheries subsidies", "subsidies contributing to overfishing",
        "sustainable use of marine resources", "small-scale fisheries",
        "increase knowledge for ocean health", "ocean health", "international maritime law",
        "blue ambition", "Coastal Climate Change Adaptation Plan", "CCCAP",
        "National Biodiversity Strategies and Action Plans", "NBSAPs", "NBSAP"
    ],

    "Ocean Finance": [
        "ocean finance", "ocean investment", "ocean-related investment",
        "ocean-related investments", "marine resource management", "financing blue carbon",
        "financing ecosystem restoration"
    ],

    "Sustainable Ocean-Based Tourism": [
        "sustainable ocean-based tourism", "sustainable coastal tourism", "coastal tourism",
        "ocean tourism", "marine pollution", "blue bonds", "blue infrastructure",
        "ocean economic sector"
    ],

"Ocean-related Storms" : [
  "tropical depression", "tropical depressions",
"tropical storm", "tropical storms","hurricane", "hurricanes",
"typhoon", "typhoons","cyclone", "cyclones","subtropical cyclone", "subtropical cyclones","extratropical cyclone", "extratropical cyclones","nor'easter", "nor'easters",
"european windstorm", "european windstorms","polar low", "polar lows","medicane", "medicanes",
"kona storm", "kona storms","mesoscale convective system", "mesoscale convective systems","squall line", "squall lines","tornadic waterspout", "tornadic waterspouts","fair-weather waterspout", "fair-weather waterspouts",

]

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