import os
import tkinter as tk
from tkinter import filedialog, messagebox
import xlrd
import re
import graphviz
from oletools.olevba import VBA_Parser

def extract_vba_from_file(file_path):
    """
    Extract VBA code from an Excel file using oletools.
    :param file_path: Path to the Excel file
    :return: List of dictionaries containing VBA filename and code
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

def analyze_vba_code(vba_code):
    """Analyzes VBA code to extract functions, subroutines, variables, and comments."""
    functions = re.findall(r'Function\s+(\w+)', vba_code, re.IGNORECASE)
    subroutines = re.findall(r'Sub\s+(\w+)', vba_code, re.IGNORECASE)
    variables = re.findall(r'Dim\s+(\w+)', vba_code, re.IGNORECASE)
    comments = re.findall(r'\'(.+)', vba_code)

    return {
        'functions': functions,
        'subroutines': subroutines,
        'variables': variables,
        'comments': comments
    }

def create_documentation(analysis_result):
    """Generates comprehensive documentation of the VBA macro analysis."""
    documentation = []

    documentation.append("## Functions:")
    for func in analysis_result['functions']:
        documentation.append(f"- {func}")

    documentation.append("\n## Subroutines:")
    for sub in analysis_result['subroutines']:
        documentation.append(f"- {sub}")

    documentation.append("\n## Variables:")
    for var in analysis_result['variables']:
        documentation.append(f"- {var}")

    documentation.append("\n## Comments:")
    for comment in analysis_result['comments']:
        documentation.append(f"- {comment}")

    return "\n".join(documentation)

def create_flowchart(analysis_result, output_file):
    """Creates a flowchart of the VBA macro process flow using Graphviz."""
    dot = graphviz.Digraph(comment='VBA Macro Process Flow')

    # Add nodes for functions and subroutines
    for func in analysis_result['functions']:
        dot.node(func, f'Function: {func}')
    for sub in analysis_result['subroutines']:
        dot.node(sub, f'Subroutine: {sub}')

    # Add edges to represent logical flow (dummy example, customize based on actual flow logic)
    for i in range(len(analysis_result['functions']) - 1):
        dot.edge(analysis_result['functions'][i], analysis_result['functions'][i + 1])
    for i in range(len(analysis_result['subroutines']) - 1):
        dot.edge(analysis_result['subroutines'][i], analysis_result['subroutines'][i + 1])

    dot.render(output_file, format='png', view=True)
    print(f"Flowchart saved to {output_file}.png")

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm")])
    if not file_path:
        return

    try:
        vba_modules = extract_vba_from_file(file_path)

        if vba_modules:
            vba_code = '\n'.join([module['code'] for module in vba_modules])
            analysis_result = analyze_vba_code(vba_code)
            documentation = create_documentation(analysis_result)
            
            output_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
            if output_path:
                with open(output_path, 'w') as doc_file:
                    doc_file.write(documentation)
                messagebox.showinfo("Success", f"Documentation saved to {output_path}")

                flowchart_output = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
                if flowchart_output:
                    create_flowchart(analysis_result, flowchart_output)
                    messagebox.showinfo("Flowchart Generated", f"Flowchart saved to {flowchart_output}.png")

        else:
            messagebox.showinfo("No VBA Code", "No VBA code found in the Excel file.")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = tk.Tk()
root.title("Upload XLSM File")
root.geometry("400x200")

# Create a label and a button
label = tk.Label(root, text="Upload XLSM File", font=("Arial", 14))
label.pack(pady=20)

upload_button = tk.Button(root, text="Upload", command=upload_file, font=("Arial", 12))
upload_button.pack(pady=20)

# Run the application
root.mainloop()
