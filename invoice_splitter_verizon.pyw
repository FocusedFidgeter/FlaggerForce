import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

# Function to split invoices
def split_invoices(pdf_file, output_dir):
    """
    Splits a PDF file containing multiple invoices into separate PDF files for each invoice.

    Parameters:
    - pdf_file (str): The path to the input PDF file.
    - output_dir (str): The directory where the output PDF files will be saved.

    Returns:
    - None

    This function takes a PDF file and an output directory as input. It reads the PDF file using the PdfReader
    class from the PyPDF2 library. It then iterates over each page of the PDF file and extracts the text from
    each page. If the text contains the string "INVOICE #:", it indicates the start of a new invoice. The function
    saves the previous invoice if there was one, and creates a new invoice PDF file using the PdfWriter class. 
    The invoice number is extracted from the text and used as the filename for the output PDF file. Finally, the
    function saves the last invoice if there was one.

    Note: This function requires the PyPDF2 library to be installed.

    Example usage:
    split_invoices("input.pdf", "output_directory")
    """
    with open(pdf_file, 'rb') as f:
        pdf_reader = PdfReader(f)
        num_pages = len(pdf_reader.pages)
        current_invoice = None
        current_writer = None

        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text = page.extract_text().strip()
            words = text.split()

            if "INVOICE #:" in text:
                if current_writer is not None:
                    # Save previous invoice
                    output_file = os.path.join(output_dir, f"{current_invoice}.pdf")
                    with open(output_file, 'wb') as output_f:
                        current_writer.write(output_f)

                # Start new invoice
                invoice_number = words[-3][3:]  # This will remove the first three characters
                current_invoice = invoice_number
                current_writer = PdfWriter()
            
            if current_writer is not None:
                current_writer.add_page(page)

        # Save the last invoice
        if current_writer is not None:
            output_file = os.path.join(output_dir, f"{current_invoice}.pdf")
            print(f"Saving {current_invoice}.pdf")
            with open(output_file, 'wb') as output_f:
                current_writer.write(output_f)

def process_files(files, output_dir):
    """
    Process the given files and split invoices for each file.

    Args:
        files (List[str]): A list of file paths to be processed.
        output_dir (str): The directory to save the split invoices.

    Returns:
        None
    """
    for file in files:
        print(f"Processing {file}")
        split_invoices(file, output_dir)

def get_directory(title):
    """
    Retrieves a directory path selected by the user.

    Args:
        title (str): The title of the directory selection dialog box.

    Returns:
        str: The selected directory path.

    Raises:
        SystemExit: If the user cancels the directory selection.

    """
    directory = filedialog.askdirectory(title=title)
    if not directory:  # If user cancels directory selection, exit the program
        exit()
    return directory

def main():
    """
    Use tkinter to select input files and output folder.
    
    Parameters:
    None
    
    Returns:
    None
    """
    # Use tkinter to select input files and output folder
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("Instructions", "This program is used to split our invoices that we download from \"My Stored Reports\" on Intaact.\n\nYou will be asked to select the files that need split and the output folder where the split invoices will be saved.\n\nA lot of nothing is going to happen, and then another message box like this will alert you when the files are all saved.")
    
    input_files = filedialog.askopenfilenames(title="Select the files with bulk invoices to split.")
    output_folder = get_directory("Select the folder where the split invoices will be moved.")

    # Process files
    process_files(input_files, output_folder)

    messagebox.showinfo("Completion", "The splitting process has been completed.\n\nYou can find the individual invoices in the folder you previously selected.")
    root.destroy()

# Run the main function
if __name__ == "__main__":
    main()