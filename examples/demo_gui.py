import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import sys
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import wmg_feedback_gen

class FileSelectorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Selector Demo")
        self.geometry("500x400")
        self.resizable(False, False)
        self.xlsx_path = tk.StringVar()
        self.docx_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.sheet_names = []
        self.selected_sheet = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        # XLSX file selection
        tk.Label(self, text="Select XLSX file containing marks:").pack(anchor="w", padx=10, pady=(10,0))
        xlsx_frame = tk.Frame(self)
        xlsx_frame.pack(fill="x", padx=10)
        tk.Entry(xlsx_frame, textvariable=self.xlsx_path, state="readonly").pack(side="left", fill="x", expand=True)
        tk.Button(xlsx_frame, text="Browse", command=self.browse_xlsx).pack(side="left", padx=5)

        # Worksheet dropdown
        tk.Label(self, text="Select Worksheet:").pack(anchor="w", padx=10, pady=(10,0))
        self.sheet_dropdown = ttk.Combobox(self, textvariable=self.selected_sheet, state="readonly", values=[])
        self.sheet_dropdown.pack(fill="x", padx=10)

        # DOCX file selection
        tk.Label(self, text="Select DOCX template file:").pack(anchor="w", padx=10, pady=(10,0))
        docx_frame = tk.Frame(self)
        docx_frame.pack(fill="x", padx=10)
        tk.Entry(docx_frame, textvariable=self.docx_path, state="readonly").pack(side="left", fill="x", expand=True)
        tk.Button(docx_frame, text="Browse", command=self.browse_docx).pack(side="left", padx=5)

        # Output directory selection
        tk.Label(self, text="Select Output Directory:").pack(anchor="w", padx=10, pady=(10,0))
        outdir_frame = tk.Frame(self)
        outdir_frame.pack(fill="x", padx=10)
        tk.Entry(outdir_frame, textvariable=self.output_dir, state="readonly").pack(side="left", fill="x", expand=True)
        tk.Button(outdir_frame, text="Browse", command=self.browse_output_dir).pack(side="left", padx=5)

        # Output filename
        tk.Label(self, text="Output Filename:").pack(anchor="w", padx=10, pady=(10,0))
        self.output_filename.set("feedback_{{STUDENTID}}.docx")
        tk.Entry(self, textvariable=self.output_filename).pack(fill="x", padx=10)

        # Generate button
        tk.Button(self, text="Generate", command=self.generate).pack(pady=20)

    def browse_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.xlsx_path.set(path)
            self.load_worksheet_names(path)

    def load_worksheet_names(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            self.sheet_names = wb.sheetnames
            self.sheet_dropdown['values'] = self.sheet_names
            if self.sheet_names:
                self.selected_sheet.set(self.sheet_names[0])
            else:
                self.selected_sheet.set('')
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load worksheet names: {e}")
            self.sheet_dropdown['values'] = []
            self.selected_sheet.set('')

    def browse_docx(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.docx_path.set(path)

    def browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def generate(self):
        # Placeholder for actual generation logic
        if not self.xlsx_path.get():
            messagebox.showwarning("Missing Input", "Please select an XLSX file.")
            return
        if not self.selected_sheet.get():
            messagebox.showwarning("Missing Input", "Please select a worksheet.")
            return
        if not self.docx_path.get():
            messagebox.showwarning("Missing Input", "Please select a DOCX file.")
            return
        if not self.output_dir.get():
            messagebox.showwarning("Missing Input", "Please select an output directory.")
            return
        if not self.output_filename.get():
            messagebox.showwarning("Missing Input", "Please enter an output filename.")
            return
        
        try:
            wmg_feedback_gen.generate(
            xlsx_filename=self.xlsx_path.get(),
            template_filename=self.docx_path.get(),
            worksheet=self.selected_sheet.get(),
            output_filename=os.path.join(self.output_dir.get(), self.output_filename.get())
            )
            messagebox.showinfo("Success", "Feedback files generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during generation:\n{e}")

if __name__ == "__main__":
    app = FileSelectorApp()
    app.mainloop()