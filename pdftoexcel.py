import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfReader

# Function to normalize headers for comparison
def normalize_header(header):
    if not header:
        return []
    return [col.strip().lower().replace("\n", " ") if col else '' for col in header]

# Function to ensure DataFrame has unique columns
def ensure_unique_columns(columns):
    """
    Ensure that the columns in a DataFrame are unique by appending a suffix to duplicates.
    """
    seen = {}
    for i, col in enumerate(columns):
        if col in seen:
            seen[col] += 1
            columns[i] = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
    return columns

# Function to clean and align DataFrame with existing combined table
def clean_and_align_dataframe(df, combined_df):
    # Ensure unique columns before processing
    df.columns = ensure_unique_columns(list(df.columns))
    
    # Find missing columns and add them with empty values
    for col in combined_df.columns:
        if col not in df.columns:
            df[col] = ''

    # Reindex to ensure columns match the combined DataFrame
    df = df[combined_df.columns]

    return df

# Function to remove blank rows and rows where the first column is empty
def remove_blank_rows(df):
    # Remove rows where all columns are NaN
    df = df.dropna(how='all')

    # Remove rows where the first column is NaN or empty
    df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0] != '')]

    return df

# Function to extract tables and save to Excel
def convert_pdf_to_excel(pdf_path, excel_path):
    print(f"PDF Path: {pdf_path}")
    print(f"Excel Path: {excel_path}")

    # Check if the directory is writable
    if not os.access(os.path.dirname(excel_path), os.W_OK):
        raise PermissionError(f"Cannot write to directory: {os.path.dirname(excel_path)}")

    # Initialize PdfReader
    reader = PdfReader(pdf_path)

    # Dictionary to store combined tables by their headers
    tables_by_header = {}

    # Extract tables from each page using pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1 and table[0]:  # Check if table has multiple rows and a valid header
                    header = normalize_header(table[0])  # Normalize header
                    header_tuple = tuple(header)  # Convert header list to tuple to use as dictionary key

                    if any(header):  # Proceed only if there's any valid header content
                        df = pd.DataFrame(table[1:], columns=header)  # Create DataFrame from the table

                        # Ensure columns are unique
                        df.columns = ensure_unique_columns(list(df.columns))

                        # Log the headers for debugging
                        print(f"Page {page_num + 1} headers: {header}")

                        if header_tuple in tables_by_header:
                            # Align the new DataFrame with the existing combined DataFrame
                            df = clean_and_align_dataframe(df, tables_by_header[header_tuple])

                            # Append the aligned DataFrame to the combined DataFrame
                            tables_by_header[header_tuple] = pd.concat([tables_by_header[header_tuple], df], ignore_index=True)
                        else:
                            # If this header is new, start a new entry in the dictionary
                            tables_by_header[header_tuple] = df
                    else:
                        print(f"Skipping a table on page {page_num + 1} due to invalid or empty header.")
                else:
                    print(f"Skipping a table on page {page_num + 1} due to missing or invalid data.")

    # Save grouped tables to Excel
    if tables_by_header:
        try:
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                for i, (header_tuple, combined_table) in enumerate(tables_by_header.items()):
                    # Ensure unique columns in the final combined DataFrame
                    combined_table.columns = ensure_unique_columns(list(combined_table.columns))

                    # Remove blank rows and rows where the first column is blank
                    combined_table = remove_blank_rows(combined_table)

                    sheet_name = f"Combined_Table_{i+1}"
                    combined_table.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Tables have been successfully extracted and saved to {excel_path}")
        except Exception as e:
            print(f"Failed to save Excel file: {e}")
            messagebox.showerror("Error", f"Failed to save Excel file: {e}")
    else:
        messagebox.showinfo("No Tables Found", "No tables were found in the PDF.")

# Function to open file dialog for PDF selection
def select_pdf():
    file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

# Function to open file dialog for Excel save location
def select_excel_save_location():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)

# Function to start conversion process
def start_conversion():
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()

    if not pdf_path or not excel_path:
        messagebox.showwarning("Input Required", "Please select both the PDF file and the destination to save the Excel file.")
    else:
        convert_pdf_to_excel(pdf_path, excel_path)

# Setup tkinter window
root = tk.Tk()
root.title("PDF to Excel Converter")

# PDF File Selection
tk.Label(root, text="Select PDF File:").grid(row=0, column=0, padx=10, pady=10)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.grid(row=0, column=1, padx=10, pady=10)
pdf_button = tk.Button(root, text="Browse", command=select_pdf)
pdf_button.grid(row=0, column=2, padx=10, pady=10)

# Excel File Save Location
tk.Label(root, text="Save Excel File As:").grid(row=1, column=0, padx=10, pady=10)
excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=1, column=1, padx=10, pady=10)
excel_button = tk.Button(root, text="Browse", command=select_excel_save_location)
excel_button.grid(row=1, column=2, padx=10, pady=10)

# Convert Button
convert_button = tk.Button(root, text="Convert", command=start_conversion, bg="green", fg="white")
convert_button.grid(row=2, columnspan=3, pady=20)

# Run the tkinter main loop
root.mainloop()
