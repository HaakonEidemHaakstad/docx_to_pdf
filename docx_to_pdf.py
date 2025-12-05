import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from docx2pdf import convert
import os
import threading
import re

class DocxToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX to PDF Converter")
        self.root.geometry("400x300")
        self.root.config(bg="#f0f0f0")

        # UI Elements
        self.frame = tk.Frame(self.root, bg="white", bd=2, relief="groove")
        self.frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.label = tk.Label(
            self.frame, 
            text="Drag & Drop .docx files here", 
            bg="white", 
            fg="#555",
            font=("Arial", 14)
        )
        self.label.pack(fill="both", expand=True)

        # Register the frame for drag and drop
        self.frame.drop_target_register(DND_FILES)
        self.frame.dnd_bind('<<Drop>>', self.drop_files)

        self.status_label = tk.Label(self.root, text="Ready", bg="#f0f0f0", anchor="w")
        self.status_label.pack(fill="x", padx=20, pady=(0, 10))

    def parse_dropped_files(self, data):
        """
        TkinterDnD returns filenames as a space-separated string.
        Paths with spaces are enclosed in curly braces {}.
        We need regex to parse this correctly.
        """
        # Regex to match paths inside curly braces OR standard paths separated by spaces
        pattern = r'\{.*?\}|\S+'
        files = re.findall(pattern, data)
        
        # Clean up the paths (remove curly braces)
        cleaned_files = [f.strip('{}') for f in files]
        return cleaned_files

    def drop_files(self, event):
        files = self.parse_dropped_files(event.data)
        docx_files = [f for f in files if f.lower().endswith('.docx')]

        if not docx_files:
            messagebox.showerror("Error", "No .docx files found!")
            return

        # Run conversion in a separate thread so the UI doesn't freeze
        threading.Thread(target=self.convert_files, args=(docx_files,)).start()

    def convert_files(self, files):
        total = len(files)
        self.update_status(f"Processing {total} file(s)...")

        for index, file_path in enumerate(files):
            try:
                self.update_status(f"Converting: {os.path.basename(file_path)}")
                
                # Convert logic
                convert(file_path)
                
            except Exception as e:
                print(f"Error converting {file_path}: {e}")
        
        self.update_status("Conversion Complete!")
        messagebox.showinfo("Success", f"Successfully converted {total} file(s)!")
        self.update_status("Ready")

    def update_status(self, text):
        self.status_label.config(text=text)

if __name__ == "__main__":
    # Note: We use TkinterDnD.Tk() instead of standard tk.Tk()
    root = TkinterDnD.Tk()
    app = DocxToPdfApp(root)
    root.mainloop()