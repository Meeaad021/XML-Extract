import lxml.etree as ET
import re
import os
import xlwt

# --- Select XML File from Current Folder ---
def choose_xml_file():
    current_dir = os.path.dirname(__file__)
    xml_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.xml')]

    if not xml_files:
        print("No XML files found in the current directory.")
        return None

    print("\nAvailable XML files:")
    for i, filename in enumerate(xml_files, start=1):
        print(f"{i}. {filename}")

    while True:
        try:
            choice = int(input("\nEnter the number of the XML file to process: "))
            if 1 <= choice <= len(xml_files):
                return os.path.join(current_dir, xml_files[choice - 1])
            else:
                print(f"Please enter a number between 1 and {len(xml_files)}.")
        except ValueError:
            print("Invalid input. Please enter a number.")

# --- Main Script ---
def main():
    xml_file_path = choose_xml_file()
    if not xml_file_path:
        return

    xml_file_name = os.path.basename(xml_file_path)
    excel_output_file = os.path.splitext(xml_file_name)[0] + ' Data.xls'

    print(f"\nAttempting to read XML from: {xml_file_path}")

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

        # Remove attributes with 'miext:' prefix
        processed_xml_string = re.sub(r'\s*miext:[a-zA-Z0-9_-]+="[^"]*"', '', processed_xml_string)

        # Remove 'miext:' prefix from tag names
        processed_xml_string = re.sub(r'<(/?)(miext:)?([^>]+)>', r'<\1\3>', processed_xml_string)

        # Escape unescaped ampersands
        processed_xml_string = re.sub(r'&(?![#a-zA-Z0-9]{2,};)', '&amp;', processed_xml_string)
        print("Escaped unescaped '&' characters.")

        # Wrap with a temporary root
        wrapped_xml_string = f"<temp_root>{processed_xml_string}</temp_root>"
        print("Wrapped XML content with <temp_root> for parsing.")

        # Parse XML
        parser = ET.XMLParser(recover=True)
        root = ET.fromstring(wrapped_xml_string.encode('utf-8'), parser=parser)
        print("Parsed XML with recovery enabled.")

        # Find all <ques> elements
        ques_elements = root.findall('.//ques')

        print(f"\n--- Preparing to write to '{excel_output_file}' ---")

        # --- Excel Writing ---
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Ques Data")

        # Headers
        sheet.write(0, 0, "Ques ID")
        sheet.write(0, 1, "Subq Filter")

        current_row = 1

        # --- Define IDs or patterns to ignore ---
        ignore_ids = {"TemplateLog", "SurveyLog", "ip_1", "REGISTRY", 
                      "S_FIELDS", "Pre_Check", "Timestamp_Start", "ExtPanelID_Opinium", 
                      "BaseURL_Opinium", "SecureKey", "keyID", "hashedURL", 
                      "ValidateDynata", "DynataQuotaFullPre", "DynataQuotaFullPath", 
                      "DynataQuotaFullHash", "DynataQuotaFull", "DynataScreenedPre", 
                      "DynataScreenedPath", "DynataScreenedHash", "DynataScreened", 
                      "DynataCompletePre", "DynataCompletePath", "DynataCompleteHash", 
                      "DynataComplete", "DynataComposite", "MailingRef", "BaseURL_IFA", 
                      "BaseURL", "Duration_Main", "TimerStart_Closing", 
                      "Duration_Closing", "ip_2", "Feedback", "Timestamp_End", 
                      "LOI", "Completed", "Screenout", "LOI_Screenout", 
                      "Quality", "LOI_Quality"}  
        
        ignore_patterns = ["demo", "intro", "hidden"]  # substrings (case-insensitive)

        if ques_elements:
            for ques_elem in ques_elements:
                ques_id = ques_elem.get('id') or ""

                # --- Skip ignored elements ---
                if (
                    ques_id in ignore_ids
                    or any(pat.lower() in ques_id.lower() for pat in ignore_patterns)
                    or ques_id.lower().startswith(("qt", "Screener"))
                ):
                    continue  # skip this ques entirely

                # --- Extract <subq> ---
                subq_element = ques_elem.find('subq')
                if subq_element is not None:
                    subq_filter = subq_element.get('filter') or ""
                else:
                    subq_filter = ""  # leave blank if none

                # --- Write row ---
                sheet.write(current_row, 0, ques_id if ques_id else "Not Found")
                sheet.write(current_row, 1, subq_filter)
                current_row += 1

            workbook.save(excel_output_file)
            print(f"✅ Successfully extracted {current_row - 1} entries to '{excel_output_file}'")

        else:
            print("⚠️ No <ques> elements found in the document. No Excel file was created.")

    except FileNotFoundError:
        print(f"❌ Error: File not found at '{xml_file_path}'.")
    except ET.XMLSyntaxError as e:
        print(f"❌ XML Syntax Error: {e}")
    except ValueError as e:
        print(f"⚠️ Data Processing Error: {e}")
    except Exception as e:
        print(f"❗ Unexpected error: {e}")

# --- Entry Point ---
if __name__ == "__main__":
    main()
