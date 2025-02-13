try:
    import os
    import pdfplumber
    import pandas as pd
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from openpyxl.styles import numbers
    from pdf2docx import Converter
    from pypdf import PdfReader, PdfWriter
    from copy import copy
except Exception as e:
    print(e)
    input()


class PDFtoExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Converter")
        self.root.geometry("550x450")
        self.root.resizable(False, False)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure("TButton", font=("Arial", 12))
        self.style.configure("TLabel", font=("Arial", 12))
        self.style.configure("RemoveButton.TButton", font=("Arial", 10), padding=2)
        
        self.pdf_files = []
        self.setup_ui()

    def setup_ui(self):
        """Centralized UI setup"""
        self.label = ttk.Label(self.root, text="Select PDF files to convert:")
        self.label.pack(pady=10)

        self.setup_list_frame()
        
        self.setup_button_frame()

    def setup_list_frame(self):
        """Setup the scrollable list frame"""
        self.list_frame = ttk.Frame(self.root)
        self.list_frame.pack(pady=5, padx=20, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.list_frame)
        self.scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_button_frame(self):
        """Setup the button frame"""
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(pady=15)

        buttons = [
            ("Add PDF(s)", self.add_pdf),
            ("Convert to Excel", self.convert_pdfs_to_excel),
            ("Convert to Word", self.convert_to_word),
            ("Merge PDFs", self.merge_pdfs),
        ]

        for text, command in buttons:
            ttk.Button(self.button_frame, text=text, command=command).pack(side=tk.LEFT, padx=5)

    def remove_all(self):
        """Remove all files from the list"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.pdf_files.clear()

    def create_file_row(self, file_path):
        """Create a new row for a file with its remove button"""
        row_frame = ttk.Frame(self.scrollable_frame)
        row_frame.pack(fill=tk.X, padx=5, pady=2)
  
        file_label = ttk.Label(row_frame, text=os.path.basename(file_path))
        file_label.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        remove_btn = ttk.Button(
            row_frame, 
            text="✕", 
            style="RemoveButton.TButton",
            command=lambda: self.remove_file(file_path, row_frame)
        )
        remove_btn.pack(side=tk.RIGHT)

    def add_pdf(self):
        """Allow user to select multiple PDF files"""
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.create_file_row(file)

    def remove_file(self, file_path, row_frame):
        """Remove a file from the list"""
        self.pdf_files.remove(file_path)
        row_frame.destroy()

    def extract_data_from_pdf(self, pdf_path):
        """Extracts table data from PDF and formats numbers for Excel."""
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)

        df = pd.DataFrame(all_data)
        df = df.replace('', None)
        
        for col in df.columns:
            pd.to_numeric(df[col], errors='ignore')
            
        return df
    
    def convert_to_word(self):
        """Convert PDFs to Word documents"""
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDF file selected!")
            return

        try:
            if len(self.pdf_files) == 1:
                docx_file_path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word Document", "*.docx")],
                    title="Save Word File As"
                )
                if docx_file_path:
                    cv = Converter(self.pdf_files[0])
                    cv.convert(docx_file_path)
                    cv.close()
                    messagebox.showinfo("Success", f"Converted and saved: {os.path.basename(docx_file_path)}")

            else:
                save_dir = filedialog.askdirectory(title="Select Folder to Save Word Files")
                if save_dir:
                    for pdf_path in self.pdf_files:
                        cv = Converter(pdf_path)
                        output_path = os.path.join(save_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".docx")
                        cv.convert(output_path)
                        cv.close()
                    messagebox.showinfo("Success", f"Converted {len(self.pdf_files)} PDFs to Word!")
            
            self.remove_all()
        except Exception as e:
            messagebox.showerror("Error", f"Error converting to Word: {str(e)}")

    def standardize_page_size(self, input_page, target_width, target_height):
        """Standardize page size while maintaining aspect ratio"""
        current_width = float(input_page.mediabox.width)
        current_height = float(input_page.mediabox.height)
        
        width_scale = target_width / current_width
        height_scale = target_height / current_height
        scale = min(width_scale, height_scale)
        
        new_page = copy(input_page)
        new_page.scale(scale, scale)
        
        return new_page
        
    def merge_pdfs(self):
        """Merges multiple PDF files into a single PDF with standardized page sizes."""
        if len(self.pdf_files) < 2:
            messagebox.showerror("Error", "Please select at least two PDF files to merge!")
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Save Merged PDF As"
            )
            
            if output_path:
                merger = PdfWriter()
                
                max_width = 0
                max_height = 0
                for pdf_path in self.pdf_files:
                    with open(pdf_path, 'rb') as file:
                        reader = PdfReader(file)
                        for page in reader.pages:
                            max_width = max(max_width, float(page.mediabox.width))
                            max_height = max(max_height, float(page.mediabox.height))

                for pdf_path in self.pdf_files:
                    with open(pdf_path, 'rb') as file:
                        reader = PdfReader(file)
                        for page in reader.pages:
                            standardized_page = self.standardize_page_size(page, max_width, max_height)
                            merger.add_page(standardized_page)

                with open(output_path, 'wb') as output_file:
                    merger.write(output_file)
                
                messagebox.showinfo("Success", f"Merged {len(self.pdf_files)} PDFs successfully!")
                self.remove_all()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error merging PDFs: {str(e)}")
        finally:
            if 'merger' in locals():
                merger.close()

    def save_to_excel_with_number_formatting(self, df, excel_path):
        """Save DataFrame to Excel with proper number formatting."""
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
            workbook = writer.book
            worksheet = workbook.active 
    
            for col in worksheet.columns:
                for cell in col:
                    if cell.row == 1:
                        continue
                    try:
                        value = float(str(cell.value).replace(',', '').replace('%', ''))
                        cell.value = value
                        
                        if value.is_integer():
                            cell.number_format = numbers.FORMAT_NUMBER
                        else:
                            cell.number_format = numbers.FORMAT_NUMBER_00
                    except (ValueError, TypeError):
                        continue

    def convert_pdfs_to_excel(self):
        """Convert selected PDFs to Excel and save the files."""
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDF file selected!")
            return

        try:
            if len(self.pdf_files) == 1:
                excel_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    title="Save Excel File As"
                )
                if excel_path:
                    df = self.extract_data_from_pdf(self.pdf_files[0])
                    self.save_to_excel_with_number_formatting(df, excel_path)
                    messagebox.showinfo("Success", f"Converted and saved: {os.path.basename(excel_path)}")

            else:
                save_dir = filedialog.askdirectory(title="Select Folder to Save Excel Files")
                if save_dir:
                    for pdf_path in self.pdf_files:
                        df = self.extract_data_from_pdf(pdf_path)
                        excel_path = os.path.join(save_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".xlsx")
                        self.save_to_excel_with_number_formatting(df, excel_path)
                    messagebox.showinfo("Success", f"Converted {len(self.pdf_files)} PDFs to Excel!")
            
            self.remove_all()
        except Exception as e:
            messagebox.showerror("Error", f"Error converting to Excel: {str(e)}")


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = PDFtoExcelConverter(root)

        root.mainloop()
    except Exception as e:
        print(e)
        input()