import os
import olefile
from oletools.olevba import VBA_Parser

def extract_vba_from_file(file_path):
    """
    Extract VBA code from an Excel file.
    :param file_path: Path to the Excel file
    :return: List of VBA code modules
    """
    vba_modules = []

    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return vba_modules

    # Open the file with olefile
    if olefile.isOleFile(file_path):
        print(f"Found OLE file: {file_path}")
        vba_parser = VBA_Parser(file_path)
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
            print(f"Extracted VBA module: {vba_filename}")
            vba_modules.append({
                'filename': vba_filename,
                'code': vba_code
            })
    else:
        print(f"Not an OLE file: {file_path}")

    return vba_modules

# Test with your file path
excel_file_path = 'C:\\Users\\Prathika\\Downloads\\kutty.xlsm'
vba_modules = extract_vba_from_file(excel_file_path)

if vba_modules:
    for module in vba_modules:
        print(f"Module: {module['filename']}")
        print(module['code'])
else:
    print("No VBA code found in the Excel file.")
