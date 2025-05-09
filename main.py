import os
from pdf2docx import Converter
from tkinter import Tk, Frame, Label, Button, filedialog, messagebox, BOTH, LEFT, RIGHT, TOP, X
import time

# Try to import tkinterdnd2 for drag and drop functionality
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    tkinterdnd2_available = True
except ImportError:
    tkinterdnd2_available = False
    print("Note: Drag and drop functionality requires tkinterdnd2. Install with: pip install tkinterdnd2")

class PDFToWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Word Converter")
        self.root.geometry("500x300")
        
        # Main frame
        self.main_frame = Frame(root)
        self.main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Drag and drop area (only if tkinterdnd2 is available)
        self.drop_frame = Frame(self.main_frame, bd=2, relief="groove")
        self.drop_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        self.drop_label = Label(
            self.drop_frame, 
            text="Drag & Drop PDF File Here\nor\nClick 'Select PDF' Button Below",
            height=10,
            font=("Arial", 12))
        self.drop_label.pack(expand=True)
        
        # Configure drag and drop if available
        if tkinterdnd2_available:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        else:
            self.drop_label.config(text="Click 'Select PDF' Button Below\n(Drag & drop not available)")
        
        # Button frame
        self.button_frame = Frame(self.main_frame)
        self.button_frame.pack(fill=X)
        
        # Select file button
        self.select_button = Button(
            self.button_frame,
            text="Select PDF File",
            command=self.select_file,
            width=15)
        self.select_button.pack(side=LEFT, padx=(0, 10))
        
        # Convert button
        self.convert_button = Button(
            self.button_frame,
            text="Convert to Word",
            command=self.convert_file,
            state="disabled",
            width=15)
        self.convert_button.pack(side=RIGHT)
        
        # Status label
        self.status_label = Label(self.main_frame, text="", fg="blue")
        self.status_label.pack(fill=X, pady=(5, 0))
        
        # File path storage
        self.pdf_path = None
        self.docx_path = None
    
    def on_drop(self, event):
        """Handle file drop event"""
        # Get the dropped file path
        file_path = event.data.strip()
        
        # On Windows, the path might be surrounded by {}
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        # Check if it's a PDF file
        if file_path.lower().endswith('.pdf'):
            self.pdf_path = file_path
            self.docx_path = os.path.splitext(file_path)[0] + ".docx"
            self.drop_label.config(text=f"Selected:\n{os.path.basename(file_path)}")
            self.convert_button.config(state="normal")
            self.status_label.config(text="Ready to convert", fg="green")
        else:
            self.status_label.config(text="Please drop a PDF file only", fg="red")
    
    def select_file(self):
        """Handle file selection via button"""
        pdf_file = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if pdf_file:
            self.pdf_path = pdf_file
            self.docx_path = os.path.splitext(pdf_file)[0] + ".docx"
            self.drop_label.config(text=f"Selected:\n{os.path.basename(pdf_file)}")
            self.convert_button.config(state="normal")
            self.status_label.config(text="Ready to convert", fg="green")
    
    def convert_file(self):
        """Convert the selected PDF to Word"""
        if not self.pdf_path:
            self.status_label.config(text="No file selected", fg="red")
            return
        
        self.status_label.config(text="Converting...", fg="blue")
        self.root.update()  # Update the GUI
        
        try:
            start_time = time.time()
            
            # Initialize Converter
            cv = Converter(self.pdf_path)
            
            # Start conversion
            cv.convert(self.docx_path, start=0, end=None)
            cv.close()
            
            end_time = time.time()
            conversion_time = end_time - start_time
            
            message = f"Conversion successful!\nTime taken: {conversion_time:.2f} seconds\nSaved to: {self.docx_path}"
            self.status_label.config(text=message, fg="green")
            messagebox.showinfo("Success", message)
            
            # Reset for next conversion
            self.convert_button.config(state="disabled")
            
        except Exception as e:
            error_msg = f"Conversion error: {str(e)}"
            self.status_label.config(text=error_msg, fg="red")
            messagebox.showerror("Error", error_msg)

def main():
    if tkinterdnd2_available:
        root = TkinterDnD.Tk()  # Use TkinterDnD's Tk if available
    else:
        root = Tk()  # Fallback to regular Tkinter
    
    app = PDFToWordConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
