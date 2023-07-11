import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFCombiner:
    def __init__(self, root):
        self.files = []
        root.geometry('600x500') # set window size

        # Set a theme color
        bg_color = '#c7d5e0' # light blue
        fg_color = '#264653' # dark blue

        root.configure(bg=bg_color)

        self.frame = tk.Frame(root, bg=bg_color)
        self.frame.pack(pady=10, fill='both', expand=True)

        self.listbox = tk.Listbox(self.frame, bg=bg_color, fg=fg_color, height=20)
        self.listbox.pack(pady=10, fill='both', expand=True)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.drop)

        self.button_frame = tk.Frame(root, bg=bg_color)
        self.button_frame.pack(pady=10)

        self.add_button = tk.Button(self.button_frame, text='Add PDFs', command=self.add_pdfs, bg=fg_color, fg=bg_color)
        self.add_button.pack(side=tk.LEFT, padx=10)

        self.remove_button = tk.Button(self.button_frame, text='Remove PDF', command=self.remove_pdf, bg=fg_color, fg=bg_color)
        self.remove_button.pack(side=tk.LEFT, padx=10)

        self.move_up_button = tk.Button(self.button_frame, text='Move Up', command=self.move_up, bg=fg_color, fg=bg_color)
        self.move_up_button.pack(side=tk.LEFT, padx=10)

        self.move_down_button = tk.Button(self.button_frame, text='Move Down', command=self.move_down, bg=fg_color, fg=bg_color)
        self.move_down_button.pack(side=tk.LEFT, padx=10)

        self.combine_button = tk.Button(self.button_frame, text='Combine PDFs', command=self.combine_pdfs, bg=fg_color, fg=bg_color)
        self.combine_button.pack(side=tk.LEFT, padx=10)

        self.split_button = tk.Button(self.button_frame, text='Split PDF', command=self.split_pdf, bg=fg_color, fg=bg_color)
        self.split_button.pack(side=tk.LEFT, padx=10)

        self.progress = Progressbar(root, length=500, mode='determinate')
        self.progress.pack(pady=10)

    def add_pdfs(self):
        files = filedialog.askopenfilenames(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if files:
            self.files.extend(files)
            for file in files:
                self.listbox.insert(tk.END, os.path.basename(file))

    def drop(self, event):
        files = root.tk.splitlist(event.data)
        valid_files = [f for f in files if f.endswith('.pdf')]
        self.files.extend(valid_files)
        for file in valid_files:
            self.listbox.insert(tk.END, os.path.basename(file))

    def remove_pdf(self):
        selected = self.listbox.curselection()
        if selected:
            self.files.pop(selected[0])
            self.listbox.delete(selected)

    def move_up(self):
        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            if index != 0:  # not already at the top
                self.files.insert(index-1, self.files.pop(index))
                self.listbox.delete(index)
                self.listbox.insert(index-1, os.path.basename(self.files[index-1]))
                self.listbox.select_set(index-1)

    def move_down(self):
        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            if index != len(self.files)-1:  # not already at the bottom
                self.files.insert(index+1, self.files.pop(index))
                self.listbox.delete(index)
                self.listbox.insert(index+1, os.path.basename(self.files[index+1]))
                self.listbox.select_set(index+1)

    def combine_pdfs(self):
        merger = PdfMerger()
        total_files = len(self.files)
        self.progress['maximum'] = total_files

        for index, pdf in enumerate(self.files):
            merger.append(pdf)
            self.progress['value'] = index + 1
            root.update_idletasks()

        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output:
            merger.write(output)
            merger.close()
            messagebox.showinfo("Success", f"PDFs combined successfully into {os.path.basename(output)}")
            self.progress['value'] = 0

    def split_pdf(self):
        file = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file:
            pdf = PdfReader(file)
            total_pages = len(pdf.pages)
            self.progress['maximum'] = total_pages

            for page in range(total_pages):
                pdf_writer = PdfWriter()
                pdf_writer.add_page(pdf.pages[page])

                output_filename = f"{os.path.splitext(os.path.basename(file))[0]}_page_{page+1}.pdf"
                output_path = os.path.join(os.path.dirname(file), output_filename)

                with open(output_path, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)

                self.progress['value'] = page + 1
                root.update_idletasks()

            messagebox.showinfo("Success", f"PDF split successfully into {total_pages} files.")
            self.progress['value'] = 0

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title('PDF Combiner')
    app = PDFCombiner(root)
    root.mainloop()
