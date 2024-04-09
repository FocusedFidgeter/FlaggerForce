# invoice_splitter.py

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

# Function to split invoices
def split_invoices(pdf_file, output_dir, mode):
    """
    Splits a PDF file into separate invoices and saves them as individual PDF files.

    Args:
        pdf_file (str): The path to the input PDF file.
        output_dir (str): The directory where the output PDF files will be saved.
        mode (str): The mode indicating the type of invoice format. Valid options are 'verizon' and 'duke'.

    Returns:
        None
    """
    with open(pdf_file, 'rb') as f:
        pdf_reader = PdfReader(f)
        num_pages = len(pdf_reader.pages)
        current_invoice = None
        current_writer = None

        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text = page.extract_text().strip()

            if "INVOICE #:" in text:
                words = text.split()
                if current_writer is not None:
                    # Save previous invoice
                    output_file = os.path.join(output_dir, f"{current_invoice}.pdf")
                    with open(output_file, 'wb') as output_f:
                        current_writer.write(output_f)

                # Start new invoice
                if mode == 'verizon':
                    invoice_number = words[-3][3:]  # Remove the 'INV' from the number.
                elif mode == 'duke':
                    invoice_number = words[-3]  # Keep the 'INV' in front of the string.
                current_invoice = invoice_number
                print(f"Processing {current_invoice}.pdf")
                current_writer = PdfWriter()
            
            if current_writer is not None:
                current_writer.add_page(page)

        # Save the last invoice
        if current_writer is not None:
            output_file = os.path.join(output_dir, f"{current_invoice}.pdf")
            with open(output_file, 'wb') as output_f:
                current_writer.write(output_f)

# Function to process files
def process_files(files, output_folder, mode):
    for file in files:
        print(f"Splitting {file}")
        split_invoices(file, output_folder, mode)

# Function to get directory
def get_directory():
    output_folder = filedialog.askdirectory(title="Select the folder where the split invoices will be moved.")
    return output_folder

# Function to open files
def open_files(root, completion_label):
    input_files = filedialog.askopenfilenames(title="Select the files with bulk invoices to split.")
    if input_files:
        completion_label.config(text="")
        root.filename = input_files

# Function to get output folder
def get_output_folder(root, completion_label):
    output_folder = get_directory()
    if output_folder:
        completion_label.config(text="")
        root.output_folder = output_folder

# Function to split files
def split_files(root, mode, completion_label):
    if hasattr(root, "filename") and hasattr(root, "output_folder"):
        process_files(root.filename, root.output_folder, mode)
        completion_label.config(text="The splitting process has been completed.\n\nYou can find the individual invoices in the folder you previously selected.")

# Main function updated to include mode selection via GUI
def main():
    """
    A function to execute the main logic of the program, which includes displaying a simple GUI for the user to
    select the files they want to split and the output folder where the split invoices will be saved, processing
    the files with the selected mode, and displaying completion message.
    """
    # Tkinter GUI setup.
    root = tk.Tk()
    root.title("Invoice Splitter")
    root.resizable(False, False)

    # Mode selection
    mode_label = tk.Label(root, text="Select the mode of operation:")
    mode_label.pack(pady=(10, 0))
    
    mode_var = tk.StringVar()
    verizon_rb = tk.Radiobutton(root, text="Verizon", variable=mode_var, value='verizon')
    duke_rb = tk.Radiobutton(root, text="Duke", variable=mode_var, value='duke')
    verizon_rb.pack()
    duke_rb.pack()

    # Function updated to get mode from radio button
    def get_mode():
        return mode_var.get()

    # Instructions label (updated to use dynamic mode display)
    def update_instructions():
        mode = get_mode()
        label.config(text=f"Splitting invoices for: {mode.upper()}")
    
    mode_var.trace_add('write', lambda *args: update_instructions())

    label = tk.Label(root, text="")
    label.pack(pady=10)

    # The rest of the GUI components remain unchanged except the calls to process_files and split_files need to pass get_mode()

    # Updated function to pass selected mode
    def split_files(root, completion_label):
        mode = get_mode()  # Get the selected mode
        if hasattr(root, "filename") and hasattr(root, "output_folder"):
            process_files(root.filename, root.output_folder, mode)
            completion_label.config(text="The splitting process has been completed.\n\nYou can find the individual invoices in the folder you previously selected.")

    # Updated button commands to not pass mode directly
    open_button = tk.Button(root, text="Downloaded Invoices", command=lambda: open_files(root, completion_label))
    open_button.pack(padx=10, pady=5)

    output_button = tk.Button(root, text="Select new invoice folder", command=lambda: get_output_folder(root, completion_label))
    output_button.pack(padx=10, pady=5)

    separator = tk.Frame(root, height=2, bg="gray")
    separator.pack(fill="x", pady=10)

    split_button = tk.Button(root, text="Split invoices", command=lambda: split_files(root, completion_label))
    split_button.pack(padx=10, pady=5)

    completion_label = tk.Label(root, text="")
    completion_label.pack()

    root.mainloop()

# Run the main function
if __name__ == "__main__":
    main()
