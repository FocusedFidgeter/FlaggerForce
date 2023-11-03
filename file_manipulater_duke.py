import os
import glob
import shutil
from tkinter import Tk, Label, Button

class FileManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Management App")
        
        # Define the source and target directories
        self.source_dir = './Outputs'
        self.target_dir = './Invoices'
        
        # Create labels
        self.label1 = Label(root, text="1. Delete CTR Files")
        self.label1.pack()
        
        self.label2 = Label(root, text="2. Move CTR Files")
        self.label2.pack()
        
        self.label3 = Label(root, text="3. Delete INV Files")
        self.label3.pack()
        
        self.label4 = Label(root, text="4. Move INV Files")
        self.label4.pack()
        
        # Create buttons
        self.button1 = Button(root, text="Delete CTR Files", command=self.delete_ctr_files)
        self.button1.pack()
        
        self.button2 = Button(root, text="Move CTR Files", command=self.move_ctr_files)
        self.button2.pack()
        
        self.button3 = Button(root, text="Delete INV Files", command=self.delete_inv_files)
        self.button3.pack()
        
        self.button4 = Button(root, text="Move INV Files", command=self.move_inv_files)
        self.button4.pack()
        
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
