import os
import glob
import shutil
from tkinter import Tk, Label, Button, Text, END

class FileManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Duke File Fixer")
        self.root.geometry("350x275")  # Adjust window size
        
        # Define the source and target directories
        self.source_dir = './Outputs'
        self.target_dir = './Invoices'
        
        # Create labels
        Label(root, text="  Delete CTR Files").grid(row=0, column=0, sticky="W")
        Label(root, text="  Move CTR Files").grid(row=2, column=0, sticky="W")
        Label(root, text="  Delete INV Files").grid(row=4, column=0, sticky="W")
        Label(root, text="  Move INV Files").grid(row=6, column=0, sticky="W")
        
        # Create buttons
        Button(root, text="Delete CTR Files", command=self.delete_ctr_files).grid(row=0, column=1, sticky="E")
        Button(root, text="Move CTR Files", command=self.move_ctr_files).grid(row=2, column=1, sticky="E")
        Button(root, text="Delete INV Files", command=self.delete_inv_files).grid(row=4, column=1, sticky="E")
        Button(root, text="Move INV Files", command=self.move_inv_files).grid(row=6, column=1, sticky="E")

        # Create description fields
        self.desc1 = Text(root, height=2, width=33)
        self.desc1.grid(row=1, column=0, columnspan=2, sticky="W")
        self.desc1.insert(END, "  Deletes all CTR files from the Invoices directory.")

        self.desc2 = Text(root, height=2, width=39)
        self.desc2.grid(row=3, column=0, columnspan=2, sticky="W")
        self.desc2.insert(END, "  Moves CTR files from \Outputs to the appropriate Invoices subdirectory.")

        self.desc3 = Text(root, height=2, width=36)
        self.desc3.grid(row=5, column=0, columnspan=2, sticky="W")
        self.desc3.insert(END, "  Deletes all Invoice PDFs from the Invoices directory.")

        self.desc4 = Text(root, height=2, width=42)
        self.desc4.grid(row=7, column=0, columnspan=2, sticky="W")
        self.desc4.insert(END, "  Moves Invoice PDFs from \Outputs to the appropriate Invoices subdirectory.")
        
        # Disable text editing
        self.desc1.config(state="disabled")
        self.desc2.config(state="disabled")
        self.desc3.config(state="disabled")
        self.desc4.config(state="disabled")
        
    def delete_ctr_files(self):
        # Define the target directory
        target_dir = './Invoices'
        
        # Get a list of all subdirectories in the target directory
        subdirs = os.listdir(target_dir)
        
        # Loop through all subdirectories
        for subdir in subdirs:
            # Define the target subdirectory
            target_subdir = os.path.join(target_dir, subdir)
            
            # Define the file pattern
            file_pattern = "CTR INV*.xlsx"
            
            # Get a list of all files in the target subdirectory that match the file pattern
            files = glob.glob(os.path.join(target_subdir, file_pattern))
            
            # Loop through all matching files
            for file in files:
                # Delete the file
                os.remove(file)
        
        print("CTR files deleted successfully.")
    
    def move_ctr_files(self):
        # Get a list of all files in the source directory
        files = os.listdir(self.source_dir)
        
        # Loop through all files
        for file in files:
            # Check if the file starts with 'CTR INV'
            if file.startswith('CTR INV'):
                # Extract the invoice number from the file name
                invoice_number = file[7:13]
                
                # Define the target subdirectory using the invoice number
                target_subdir = os.path.join(self.target_dir, 'INV' + invoice_number)
                
                # If the target subdirectory doesn't exist, create it
                if not os.path.exists(target_subdir):
                    os.makedirs(target_subdir)
                
                # Define the source and target file paths
                source_file = os.path.join(self.source_dir, file)
                target_file = os.path.join(target_subdir, file)
                
                # Move the file to the target subdirectory
                shutil.move(source_file, target_file)
        
        print("CTR files moved successfully.")
    
    def delete_inv_files(self):
        # Define the target directory
        target_dir = './Invoices'
        
        # Get a list of all subdirectories in the target directory
        subdirs = os.listdir(target_dir)
        
        # Loop through all subdirectories
        for subdir in subdirs:
            # Define the target subdirectory
            target_subdir = os.path.join(target_dir, subdir)
            
            # Define the file pattern
            file_pattern = "INV*.pdf"
            
            # Get a list of all files in the target subdirectory that match the file pattern
            files = glob.glob(os.path.join(target_subdir, file_pattern))
            
            # Loop through all matching files
            for file in files:
                # Delete the file
                os.remove(file)
        
        print("INV files deleted successfully.")
    
    def move_inv_files(self):
        # Get a list of all files in the source directory
        files = os.listdir(self.source_dir)
        
        # Loop through all files
        for file in files:
            # Check if the file starts with 'INV'
            if file.startswith('INV'):
                # Extract the invoice number from the file name
                invoice_number = file[3:9]
                
                # Define the target subdirectory using the invoice number
                target_subdir = os.path.join(self.target_dir, 'INV' + invoice_number)
                
                # If the target subdirectory doesn't exist, create it
                if not os.path.exists(target_subdir):
                    os.makedirs(target_subdir)
                
                # Define the source and target file paths
                source_file = os.path.join(self.source_dir, file)
                target_file = os.path.join(target_subdir, file)
                
                # Move the file to the target subdirectory
                shutil.move(source_file, target_file)
        
        print("INV files moved successfully.")

# Create the Tkinter application window
root = Tk()

# Create an instance of the FileManagementApp
app = FileManagementApp(root)

# Run the Tkinter event loop
root.mainloop()
