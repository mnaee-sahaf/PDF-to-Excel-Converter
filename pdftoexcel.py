import os
import tkinter
import tkinter as tk
import pdfplumber
import pandas as pd
import logging
import threading
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfReader

# Set up logging
logging.basicConfig(level=logging.INFO)

# Function to normalize headers for comparison
def normalize_header(header: list[str]) -> list[str]:
    """
    Normalize the headers by stripping whitespace, converting to lowercase, and replacing newline characters with spaces.

    Args:
        header (list[str]): List of header strings.

    Returns:
        list[str]: Normalized list of header strings.
    """
    if not header:
        return []
    return [col.strip().lower().replace("\n", " ") if col else '' for col in header]

# Function to ensure DataFrame has unique columns
def ensure_unique_columns(columns: list[str]) -> list[str]:
    """
    Ensure all column names are unique by appending a suffix to duplicates.

    Args:
        columns (list[str]): List of column names.

    Returns:
        list[str]: List of unique column names.
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
def clean_and_align_dataframe(df: pd.DataFrame, combined_df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean and align a DataFrame with an existing combined DataFrame.

    Args:
        df (pd.DataFrame): DataFrame to clean and align.
        combined_df (pd.DataFrame): DataFrame to align with.

    Returns:
        pd.DataFrame: Aligned DataFrame.
    """
    df.columns = ensure_unique_columns(normalize_header(list(df.columns)))
    for col in combined_df.columns:
        if col not in df.columns:
            df[col] = ''
    df = df[combined_df.columns]
    return df

# Function to remove blank rows and rows where the first column is empty
def remove_blank_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove rows that are completely blank or where the first column is empty.

    Args:
        df (pd.DataFrame): DataFrame to clean.

    Returns:
        pd.DataFrame: Cleaned DataFrame.
    """
    df = df.dropna(how='all')
    df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0] != '')]
    return df

# Function to process DataFrame
def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Process a DataFrame by normalizing headers and ensuring unique column names.

    Args:
        df (pd.DataFrame): DataFrame to process.

    Returns:
        pd.DataFrame: Processed DataFrame.
    """
    df.columns = ensure_unique_columns(normalize_header(list(df.columns)))
    return df

# Function to extract tables and save to Excel
def convert_pdf_to_excel(pdf_path: str, excel_path: str):
    """
    Extract tables from a PDF file and save them to an Excel file.

    Args:
        pdf_path (str): Path to the PDF file.
        excel_path (str): Path to save the Excel file.

    Raises:
        ValueError: If the provided paths are not valid.
    """
    if not pdf_path.endswith('.pdf'):
        raise ValueError("The provided file is not a PDF.")
    if not excel_path.endswith('.xlsx'):
        raise ValueError("The output file must have an .xlsx extension.")

    logging.info(f"PDF Path: {pdf_path}")
    logging.info(f"Excel Path: {excel_path}")

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

                        # Process DataFrame
                        df = process_dataframe(df)

                        # Log the headers for debugging
                        logging.info(f"Page {page_num + 1} headers: {header}")

                        if header_tuple in tables_by_header:
                            # Align the new DataFrame with the existing combined DataFrame
                            df = clean_and_align_dataframe(df, tables_by_header[header_tuple])

                            # Append the aligned DataFrame to the combined DataFrame
                            tables_by_header[header_tuple] = pd.concat([tables_by_header[header_tuple], df], ignore_index=True)
                        else:
                            # If this header is new, start a new entry in the dictionary
                            tables_by_header[header_tuple] = df
                    else:
                        logging.warning(f"Skipping a table on page {page_num + 1} due to invalid or empty header.")
                else:
                    logging.warning(f"Skipping a table on page {page_num + 1} due to missing or invalid data.")

    # Save grouped tables to Excel
    if tables_by_header:
        try:
            combined_tables = []
            for header_tuple, combined_table in tables_by_header.items():
                combined_table = remove_blank_rows(combined_table)
                combined_tables.append(combined_table)

            final_combined_table = pd.concat(combined_tables, ignore_index=True)

            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                final_combined_table.to_excel(writer, sheet_name="Combined_Table", index=False)

                # Adjust column widths
                worksheet = writer.sheets["Combined_Table"]
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)  # Adding some padding
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

                # Optionally adjust row heights (if needed)
                for row in worksheet.iter_rows():
                    max_height = 0
                    for cell in row:
                        if cell.value:
                            max_height = max(max_height, len(str(cell.value).splitlines()))
                    if max_height > 1:  # Set a minimum height if there are multiple lines
                        worksheet.row_dimensions[row[0].row].height = max_height * 15  # Adjust height based on content

            messagebox.showinfo("Success", f"Tables have been successfully extracted and saved to {excel_path}")
        except Exception as e:
            logging.error(f"Failed to save Excel file: {e}")
            messagebox.showerror("Error", f"Failed to save Excel file: {e}")
    else:
        messagebox.showinfo("No Tables Found", "No tables were found in the PDF.")

# Function to open file dialog for PDF selection
def select_pdf():
    """Open a file dialog to select a PDF file."""
    file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

# Function to open file dialog for Excel save location
def select_excel_save_location():
    """Open a file dialog to select a location to save the Excel file."""
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)

# Function to start conversion process
def start_conversion():
    """Retrieve paths and start the conversion process in a separate thread."""
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()
    if not pdf_path or not excel_path:
        messagebox.showwarning("Input Required", "Please select both the PDF file and the destination to save the Excel file.")
    else:
        threading.Thread(target=convert_pdf_to_excel, args=(pdf_path, excel_path)).start()

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
