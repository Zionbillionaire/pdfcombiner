import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Style
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import datetime
import webbrowser
from pdf2docx import Converter
import pandas as pd
import pdfplumber

class PDFCombiner:
    def __init__(self, root, expiry_date):
        self.files = []
        self.expiry_date = expiry_date
        root.geometry('600x500')

        bg_color = '#ADD8E6'  # Light Blue
        fg_color = '#00008B'  # Dark Blue

        root.configure(bg=bg_color)

        self.frame = tk.Frame(root, bg=bg_color)
        self.frame.pack(pady=10, fill='both', expand=True)

        self.listbox = tk.Listbox(self.frame, bg=bg_color, fg=fg_color, height=10)
        self.listbox.pack(pady=10, fill='both', expand=True)

        self.result_box = tk.Text(root, bg='white', fg=fg_color, height=6)
        self.result_box.pack(pady=10, fill='both', expand=True)

        self.button_frame1 = tk.Frame(root, bg=bg_color)
        self.button_frame1.pack(pady=10)

        self.button_frame2 = tk.Frame(root, bg=bg_color)
        self.button_frame2.pack(pady=10)

        # First Row Buttons
        self.add_button = tk.Button(self.button_frame1, text='Add PDFs', command=self.add_pdfs, bg=fg_color, fg=bg_color)
        self.add_button.pack(side=tk.LEFT, padx=10)

        self.remove_button = tk.Button(self.button_frame1, text='Remove PDF', command=self.remove_pdf, bg=fg_color, fg=bg_color)
        self.remove_button.pack(side=tk.LEFT, padx=10)

        self.move_up_button = tk.Button(self.button_frame1, text='Move Up', command=self.move_up, bg=fg_color, fg=bg_color)
        self.move_up_button.pack(side=tk.LEFT, padx=10)

        self.move_down_button = tk.Button(self.button_frame1, text='Move Down', command=self.move_down, bg=fg_color, fg=bg_color)
        self.move_down_button.pack(side=tk.LEFT, padx=10)

        # Second Row Buttons
        self.combine_button = tk.Button(self.button_frame2, text='Combine PDFs', command=self.combine_pdfs, bg=fg_color, fg=bg_color)
        self.combine_button.pack(side=tk.LEFT, padx=10)

        self.split_button = tk.Button(self.button_frame2, text='Split PDF', command=self.split_pdf, bg=fg_color, fg=bg_color)
        self.split_button.pack(side=tk.LEFT, padx=10)

        self.convert_to_word_button = tk.Button(self.button_frame2, text='Convert PDF to Word', command=self.convert_pdf_to_word, bg=fg_color, fg=bg_color)
        self.convert_to_word_button.pack(side=tk.LEFT, padx=10)

        self.convert_to_excel_button = tk.Button(self.button_frame2, text='Convert PDF to Excel', command=self.convert_pdf_to_excel, bg=fg_color, fg=bg_color)
        self.convert_to_excel_button.pack(side=tk.LEFT, padx=10)

        self.progress = Progressbar(root, length=500, mode='determinate')
        self.progress.pack(pady=10)

        self.footer_label = tk.Label(root, text=self.get_footer_text(), bg=bg_color, fg=fg_color, cursor="hand2")
        self.footer_label.pack()

        self.footer_label.bind("<Button-1>", self.open_indiangpt_profile)

        # ... rest of your code ...

    # Your methods here...

    def add_pdfs(self):
        if not self.is_key_valid():
            return

        try:
            files = filedialog.askopenfilenames(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if files:
                self.files.extend(files)
                for file in files:
                    self.listbox.insert(tk.END, os.path.basename(file))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add PDFs. Error: {str(e)}")

    def remove_pdf(self):
        if not self.is_key_valid():
            return

        selected = self.listbox.curselection()
        if selected:
            self.files.pop(selected[0])
            self.listbox.delete(selected)

    def move_up(self):
        if not self.is_key_valid():
            return

        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            if index != 0:
                self.files.insert(index-1, self.files.pop(index))
                self.listbox.delete(index)
                self.listbox.insert(index-1, os.path.basename(self.files[index-1]))
                self.listbox.select_set(index-1)

    def move_down(self):
        if not self.is_key_valid():
            return

        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            if index != len(self.files) - 1:
                self.files.insert(index+1, self.files.pop(index))
                self.listbox.delete(index)
                self.listbox.insert(index+1, os.path.basename(self.files[index+1]))
                self.listbox.select_set(index+1)

    def combine_pdfs(self):
        if not self.is_key_valid():
            return

        if self.files:
            output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if output_file:
                merger = PdfMerger()
                for file in self.files:
                    merger.append(file)
                merger.write(output_file)
                merger.close()
                messagebox.showinfo("Success", "PDFs combined successfully!")
                self.result_box.insert(tk.END, f"PDFs combined successfully to {os.path.basename(output_file)}\n")

    def split_pdf(self):
        if not self.is_key_valid():
            return

        file = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file:
            output_folder = filedialog.askdirectory()
            if output_folder:
                pdf = PdfReader(file)
                for i, page in enumerate(pdf.pages):
                    writer = PdfWriter()
                    writer.add_page(page)
                    output_file = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(file))[0]}_{i+1}.pdf")
                    writer.write(output_file)
                pdf.close()
                messagebox.showinfo("Success", "PDF split successfully!")
                self.result_box.insert(tk.END, f"PDF split successfully to {output_folder}\n")

    def convert_pdf_to_word(self):
        if not self.is_key_valid():
            return

        file = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file:
            output = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if output:
                cv = Converter(file)
                cv.convert(output, start=0, end=None)
                cv.close()
                messagebox.showinfo("Success", "PDF converted successfully to Word!")
                self.result_box.insert(tk.END, f"PDF converted successfully to {os.path.basename(output)} Word file\n")

    def convert_pdf_to_excel(self):
        if not self.is_key_valid():
            return

        try:
            file = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if file:
                output = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                if output:
                    with pdfplumber.open(file) as pdf:
                        tables = []
                        self.progress['maximum'] = len(pdf.pages)
                        self.progress['value'] = 0
                        for i, page in enumerate(pdf.pages):
                            table = page.extract_table()
                            tables.append(pd.DataFrame(table[1:], columns=table[0]))
                            self.progress['value'] += 1
                            root.update_idletasks()

                    with pd.ExcelWriter(output) as writer:
                        for i, df in enumerate(tables):
                            df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)

                    messagebox.showinfo("Success", f"PDF converted successfully to {os.path.basename(output)} Excel file")
                    self.result_box.insert(tk.END, f"PDF converted successfully to {os.path.basename(output)} Excel file\n")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert PDF to Excel. Error: {str(e)}")

    def is_key_valid(self):
        if datetime.datetime.now() > self.expiry_date:
            messagebox.showerror("Error", "Your key has expired. Please contact IndianGPT for a new key.")
            return False
        return True

    def get_footer_text(self):
        remaining_days = (self.expiry_date - datetime.datetime.now()).days
        if remaining_days < 0:
            return "Your key has expired. Please contact IndianGPT for a new key."
        return f"Key valid for next {remaining_days} days. Click here to extend."

    def open_indiangpt_profile(self, event):
        webbrowser.open("https://in.linkedin.com/in/megnush")

if __name__ == "__main__":
    root = tk.Tk()
    root.title('PDF Combiner')
    expiry_date = datetime.datetime.strptime("2023-11-30", "%Y-%m-%d")
    PDFCombiner(root, expiry_date)
    root.mainloop()
