# XML to Excel Exporter

A Python utility to extract <ques> elements from XML files and export their IDs and sub-question filter attributes to Excel. Supports ignoring specific elements based on ID, substring, or prefix.

## Table of Contents

- Clone the Repository
- Install Dependencies
- Configuration
- Running the Script
- System Overview
- Excel Output
- Error Handling
- License

## Clone the Repository

Clone the repo to your local machine:

git clone https://github.com/your-username/xml-to-excel.git
cd xml-to-excel

## Install Dependencies

Ensure you have Python 3.8+ installed.

Install required packages:

pip install lxml xlwt

- lxml → for robust XML parsing, including malformed XML recovery.
- xlwt → to write Excel .xls files.

## Configuration

### Ignore Lists

You can configure which <ques> elements to skip in the script:

ignore_ids = {"Q999", "TEST_1"}            # Exact IDs to ignore
ignore_patterns = ["demo", "intro"]       # Ignore if ID contains any substring
ignore_prefixes = ("qt", "screener")      # Ignore if ID starts with any of these

- Non-existent IDs are ignored safely.
- Prefix checks are case-insensitive.

### XML File Selection

- Place XML files in the same folder as the script.
- The script prompts you to choose which XML file to process.

## Running the Script

Run the script from the terminal:

python xml_to_excel.py

1. Select the XML file from the numbered list.
2. The script parses the XML, applies ignore rules, and writes an Excel file.
3. Output file is saved as:

<XML File Name> Data.xls

## System Overview

### XML Parsing

- Wraps XML in a temporary root for robust parsing.
- Cleans unwanted miext: prefixes and escapes unescaped &.
- Uses lxml recovery mode to handle malformed XML.

### Filtering

- Skips <ques> elements based on exact IDs, substrings, or prefixes.

## Excel Export

- Writes <ques> ID in column A.
- Writes <subq> filter attribute (or blank if missing) in column B.

### Excel Output Example

Ques ID    Subq Filter
Q101       10-20
Q102       
Q103       5-15

- Blank Subq Filter → either no <subq> element or missing filter attribute.

## Error Handling

- File not found → prints a clear error message.
- Malformed XML → recovers using lxml parser.
- Unexpected exceptions → printed with stack trace.

## License

MIT License © Meeaad
