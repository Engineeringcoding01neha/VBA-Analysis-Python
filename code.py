import os
from tkinter import Tk, Label, Button, filedialog, messagebox
from oletools.olevba import VBA_Parser

def extract_vba_from_file(file_path):
    """
    Extract VBA code from an Excel file.
    :param file_path: Path to the Excel file
    :return: List of VBA code modules
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    
    if not file_path.lower().endswith(('.xls', '.xlsm')):
        raise ValueError("Unsupported file format. Please provide an .xls or .xlsm file.")
    
    vba_modules = []
    vba_parser = VBA_Parser(file_path)

    if vba_parser.detect_vba_macros():
        print("VBA macros detected in the file.")
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
            vba_modules.append({
                'filename': vba_filename,
                'code': vba_code
            })
            print(f"Extracted VBA module: {vba_filename}")
    else:
        print("No VBA macros found in the file.")
    
    return vba_modules

def document_vba_code(vba_modules):
    """
    Generate documentation for VBA code.
    :param vba_modules: List of VBA code modules
    :return: String documentation
    """
    documentation = "VBA Code Documentation\n\n"
    for module in vba_modules:
        documentation += f"Module: {module['filename']}\n"
        documentation += "-" * len(f"Module: {module['filename']}\n")
        documentation += f"\n{module['code']}\n"
        documentation += "\n" + "=" * 50 + "\n"

    return documentation

def save_documentation(documentation, output_path):
    """
    Save documentation to a file.
    :param documentation: String documentation
    :param output_path: Path to save the documentation
    """
    with open(output_path, 'w') as doc_file:
        doc_file.write(documentation)

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm")])
    if not file_path:
        return

    try:
        vba_modules = extract_vba_from_file(file_path)

        if vba_modules:
            documentation = document_vba_code(vba_modules)
            output_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
            if output_path:
                save_documentation(documentation, output_path)
                messagebox.showinfo("Success", f"Documentation saved to {output_path}")
        else:
            messagebox.showinfo("No VBA Code", "No VBA code found in the Excel file.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = Tk()
root.title("Upload XLSM File")
root.geometry("400x200")

# Create a label and a button
label = Label(root, text="Upload XLSM File", font=("Arial", 14))
label.pack(pady=20)

upload_button = Button(root, text="Upload", command=upload_file, font=("Arial", 12))
upload_button.pack(pady=20)

# Run the application
root.mainloop()