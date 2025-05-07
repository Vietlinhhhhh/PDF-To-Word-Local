import os
from pdf2docx import Converter
from tkinter import Tk, filedialog, messagebox
import time

def convert_pdf_to_word(pdf_path, docx_path):
    """
    Convert PDF file to Word DOCX
    :param pdf_path: Input PDF file path
    :param docx_path: Output Word file path
    :return: None
    """
    try:
        # Initialize Converter
        cv = Converter(pdf_path)
        
        # Start conversion
        start_time = time.time()
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        
        end_time = time.time()
        conversion_time = end_time - start_time
        
        return True, f"Conversion successful! Time taken: {conversion_time:.2f} seconds"
    except Exception as e:
        return False, f"Conversion error: {str(e)}"

def select_file():
    root = Tk()
    root.withdraw()  # Hide the main tkinter window
    
    # Select PDF file
    pdf_file = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )
    
    if not pdf_file:
        return None, None
    
    # Suggest output filename
    output_dir = os.path.dirname(pdf_file)
    base_name = os.path.splitext(os.path.basename(pdf_file))[0]
    docx_file = os.path.join(output_dir, f"{base_name}.docx")
    
    # Let user choose where to save Word file
    docx_file = filedialog.asksaveasfilename(
        title="Save Word File",
        initialdir=output_dir,
        initialfile=f"{base_name}.docx",
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )
    
    return pdf_file, docx_file

def main():
    print("=== PDF to Word Converter ===")
    
    # Select files through GUI
    pdf_path, docx_path = select_file()
    
    if not pdf_path or not docx_path:
        print("No file selected. Exiting program.")
        return
    
    print(f"Converting: {pdf_path} -> {docx_path}")
    
    # Perform conversion
    success, message = convert_pdf_to_word(pdf_path, docx_path)
    
    # Show results
    if success:
        print(message)
        messagebox.showinfo("Success", message)
    else:
        print(message)
        messagebox.showerror("Error", message)

if __name__ == "__main__":
    main()