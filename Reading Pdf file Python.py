import os
import tkinter as tk
from tkinter import filedialog
import openpyxl

def find_pdfs(directory):
    pdf_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.pdf') or file.endswith('PDF'):
                pdf_files.append(os.path.join(root, file))
    return pdf_files


def remove_duplicates(pdf_files):
    unique_files = []
    unique_names = set()
    for pdf_file in pdf_files:
        file_name = pdf_file
        if file_name not in unique_names:
            unique_files.append(pdf_file)
            unique_names.add(file_name)
        else:
            print(f"Duplicate found: {file_name}")
            
    return unique_files

def export_to_excel(unique_pdf_files):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Unique PDF Files"
    
    # Write headers or heading on A1 row
    sheet['A1'] = "File Name"
    
    # Write data
    for row_idx, file_name in enumerate(unique_pdf_files, start=2):
        sheet[f'A{row_idx}'] = file_name
    
    # Save the workbook
    excel_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if excel_file:
        workbook.save(excel_file)
        print(f"Excel file saved successfully: {excel_file}")

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    folder_path = filedialog.askdirectory(title="Select Folder containing PDF files")
    if folder_path:
        pdf_files = find_pdfs(folder_path)
        pdf_files_with_names = [os.path.basename(pdf_path) for pdf_path in pdf_files]  # Get list of (file_name, text) tuples
        unique_pdf_files = remove_duplicates(pdf_files_with_names)
        
        if unique_pdf_files:
            export_to_excel(unique_pdf_files)
        else:
            print("No unique PDF files found.")

if __name__ == "__main__":
    main()
