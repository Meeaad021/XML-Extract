import lxml.etree as ET
import re
import os
import xlwt # Import the xlwt library for .xls files

# --- Configuration ---
xml_file_name = 'Otex Script.xml'
xml_file_path = os.path.join(os.path.dirname(__file__), xml_file_name)
excel_output_file = 'Otex Data.xls' # Name of the Excel file to create

print(f"Attempting to read XML from: {xml_file_path}")

try:
    with open(xml_file_path, 'r', encoding='utf-8-sig') as f:
        xml_source_string = f.read()

    # --- Pre-processing for robust parsing ---
    processed_xml_string = xml_source_string.lstrip()

    first_lt_index = processed_xml_string.find('<')
    if first_lt_index != -1:
        processed_xml_string = processed_xml_string[first_lt_index:]
    else:
        raise ValueError("No XML content found after initial cleanup. File might be empty or malformed.")

    # Remove ALL attributes that have the 'miext:' prefix.
    processed_xml_string = re.sub(r'\s*miext:[a-zA-Z0-9_-]+="[^"]*"', '', processed_xml_string)

    # Strip the 'miext:' prefix from element (tag) names.
    processed_xml_string = re.sub(r'<(/?)(miext:)?([^>]+)>', r'<\1\3>', processed_xml_string)

    # Escape unescaped ampersands (&).
    processed_xml_string = re.sub(r'&(?![#a-zA-Z0-9]{2,};)', '&amp;', processed_xml_string)
    print("Escaped unescaped '&' characters.")


    # --- Wrap the entire content with a temporary root ---
    wrapped_xml_string = f"<temp_root>{processed_xml_string}</temp_root>"
    print("Wrapped XML content with <temp_root> for parsing.")


    # --- Parsing and Extraction ---
    parser = ET.XMLParser(recover=True)
    root = ET.fromstring(wrapped_xml_string.encode('utf-8'), parser=parser)
    print("Parsed XML with recovery enabled.")

    # Find ALL <ques> elements anywhere in the document.
    ques_elements = root.findall('.//ques')

    print(f"\n--- Preparing to write to '{excel_output_file}' ---")

    # --- Excel Writing Setup ---
    workbook = xlwt.Workbook() # Create a new workbook
    sheet = workbook.add_sheet("Ques Data") # Add a sheet named "Ques Data"

    # Write headers to the first row (row 0)
    sheet.write(0, 0, "Ques ID")
    sheet.write(0, 1, "Subq Filter")

    current_row = 1 # Start writing data from the second row (row 1)

    # --- Extraction and Excel Writing Loop ---
    if ques_elements:
        for i, ques_elem in enumerate(ques_elements):
            ques_id = ques_elem.get('id')
            
            subq_element = ques_elem.find('subq')
            
            subq_filter = "N/A" 

            if subq_element is not None:
                subq_filter = subq_element.get('filter')
                if subq_filter is None:
                    subq_filter = "Attribute 'filter' not found in <subq> element."
            else:
                subq_filter = "No <subq> element found."
            
            # Write the extracted data to the Excel sheet
            sheet.write(current_row, 0, ques_id if ques_id else "Not Found")
            sheet.write(current_row, 1, subq_filter)
            
            current_row += 1 # Move to the next row for the next entry

        # Save the workbook after the loop finishes
        workbook.save(excel_output_file)
        print(f"Successfully extracted {current_row - 1} entries to '{excel_output_file}'")

    else:
        print("No <ques> elements found in the document. No Excel file was created.")

except FileNotFoundError:
    print(f"Error: The XML file '{xml_file_name}' was not found at '{xml_file_path}'.")
    print("Please ensure your XML file is in the same directory as this Python script and named correctly.")
except ET.XMLSyntaxError as e:
    print(f"Error parsing XML with lxml: {e}")
    print("This indicates a fundamental XML structure problem, even with recovery and extensive cleanup.")
except ValueError as e:
    print(f"Data Processing Error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")