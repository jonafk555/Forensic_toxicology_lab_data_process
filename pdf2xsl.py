import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import pandas as pd
import openpyxl

class PDFtoExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        
        self.label = tk.Label(root, text="Select a PDF file to convert:")
        self.label.pack(pady=10)
        
        self.select_button = tk.Button(root, text="Select PDF", command=self.select_pdf)
        self.select_button.pack(pady=5)
        
        self.convert_button = tk.Button(root, text="Convert to Excel", command=self.convert_to_excel, state=tk.DISABLED)
        self.convert_button.pack(pady=5)
        
        self.pdf_path = None

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            self.convert_button.config(state=tk.NORMAL)
            messagebox.showinfo("Selected File", f"Selected PDF file: {self.pdf_path}")
        else:
            self.convert_button.config(state=tk.DISABLED)
            messagebox.showwarning("No File Selected", "Please select a PDF file to convert.")
    
    def convert_to_excel(self):
        if not self.pdf_path:
            messagebox.showwarning("No File Selected", "Please select a PDF file to convert.")
            return
        
        try:
            # Extract text from PDF
            doc = fitz.open(self.pdf_path)
            text = ""
            for page in doc:
                text += page.get_text()
            
            # Convert text to DataFrame (simple split by lines for this example)
            lines = text.split('\n')
            data = {'Content': lines}
            df = pd.DataFrame(data)
            
            # Save DataFrame to Excel
            excel_path = self.pdf_path.replace('.pdf', '.xlsx')
            df.to_excel(excel_path, index=False)
            
            messagebox.showinfo("Success", f"PDF converted to Excel successfully: {excel_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoExcelConverter(root)
    root.mainloop()
