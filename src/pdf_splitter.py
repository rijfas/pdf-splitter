import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
from PyPDF2 import PdfReader, PdfWriter
import os
from threading import Thread
import webbrowser

class PDFSplitterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Extractor")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        
        self.style = ttk.Style()
        self.style.configure('Modern.TButton', padding=10, font=('Helvetica', 10))
        self.style.configure('Modern.TLabel', font=('Helvetica', 10))
        self.style.configure('Title.TLabel', font=('Helvetica', 12, 'bold'))
        self.style.configure('Link.TLabel', font=('Helvetica', 10, 'underline'), foreground='blue')
        self.style.configure('Modern.TFrame', padding=20)
        
        
        self.create_menu()
        
        
        self.main_frame = ttk.Frame(root, style='Modern.TFrame')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        
        title_label = ttk.Label(
            self.main_frame, 
            text="PDF Page Range Extractor",
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 20))
        
        
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_frame.pack(fill=tk.X, pady=5)
        
        self.input_path = tk.StringVar()
        self.input_path.trace('w', self.update_output_path)  
        
        ttk.Label(
            self.input_frame,
            text="Input PDF:",
            style='Modern.TLabel'
        ).pack(side=tk.LEFT)
        
        self.input_entry = ttk.Entry(
            self.input_frame,
            textvariable=self.input_path,
            width=50
        )
        self.input_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            self.input_frame,
            text="Browse",
            style='Modern.TButton',
            command=self.browse_input
        ).pack(side=tk.LEFT)
        
        
        self.output_frame = ttk.Frame(self.main_frame)
        self.output_frame.pack(fill=tk.X, pady=5)
        
        self.output_path = tk.StringVar()
        ttk.Label(
            self.output_frame,
            text="Output PDF:",
            style='Modern.TLabel'
        ).pack(side=tk.LEFT)
        
        self.output_entry = ttk.Entry(
            self.output_frame,
            textvariable=self.output_path,
            width=50,
            state='readonly'  
        )
        self.output_entry.pack(side=tk.LEFT, padx=5)
        
        
        self.page_frame = ttk.Frame(self.main_frame)
        self.page_frame.pack(fill=tk.X, pady=20)
        
        
        self.start_page_frame = ttk.Frame(self.page_frame)
        self.start_page_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(
            self.start_page_frame,
            text="Start Page:",
            style='Modern.TLabel'
        ).pack()
        
        self.start_page = tk.StringVar(value="1")
        self.start_page.trace('w', self.update_output_path)  
        self.start_spinbox = ttk.Spinbox(
            self.start_page_frame,
            from_=1,
            to=9999,
            textvariable=self.start_page,
            width=10
        )
        self.start_spinbox.pack()
        
        
        self.end_page_frame = ttk.Frame(self.page_frame)
        self.end_page_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(
            self.end_page_frame,
            text="End Page:",
            style='Modern.TLabel'
        ).pack()
        
        self.end_page = tk.StringVar(value="1")
        self.end_page.trace('w', self.update_output_path)  
        self.end_spinbox = ttk.Spinbox(
            self.end_page_frame,
            from_=1,
            to=9999,
            textvariable=self.end_page,
            width=10
        )
        self.end_spinbox.pack()
        
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=20)
        
        
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(
            self.main_frame,
            textvariable=self.status_var,
            style='Modern.TLabel'
        )
        self.status_label.pack(pady=5)
        
        
        self.extract_button = ttk.Button(
            self.main_frame,
            text="Extract Pages",
            style='Modern.TButton',
            command=self.start_extraction
        )
        self.extract_button.pack(pady=10)
        
        
        self.create_about_section()

    def update_output_path(self, *args):
        """Update output path based on input file and page range"""
        if self.input_path.get():
            try:
                
                input_dir = os.path.dirname(self.input_path.get())
                input_filename = os.path.splitext(os.path.basename(self.input_path.get()))[0]
                
                
                output_filename = f"{input_filename}_p{self.start_page.get()}-{self.end_page.get()}.pdf"
                
                
                output_path = os.path.join(input_dir, output_filename)
                
                self.output_path.set(output_path)
            except:
                
                self.output_path.set("")
        else:
            self.output_path.set("")

    def browse_input(self):
        filename = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if filename:
            self.input_path.set(filename)
            
            try:
                reader = PdfReader(filename)
                total_pages = len(reader.pages)
                self.start_spinbox.configure(to=total_pages)
                self.end_spinbox.configure(to=total_pages)
                self.end_page.set(str(total_pages))
            except Exception as e:
                messagebox.showerror("Error", f"Error reading PDF: {str(e)}")

    def validate_inputs(self):
        if not self.input_path.get():
            messagebox.showerror("Error", "Please select an input PDF file")
            return False
            
        try:
            start = int(self.start_page.get())
            end = int(self.end_page.get())
            
            if start < 1 or end < 1:
                raise ValueError("Page numbers must be positive")
                
            if start > end:
                raise ValueError("Start page must be less than or equal to end page")
                
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return False
            
        return True

    def extract_pages(self):
        try:
            
            reader = PdfReader(self.input_path.get())
            writer = PdfWriter()
            
            
            total_pages = len(reader.pages)
            start = int(self.start_page.get())
            end = min(int(self.end_page.get()), total_pages)
            
            
            for i in range(start - 1, end):
                self.status_var.set(f"Extracting page {i + 1}...")
                writer.add_page(reader.pages[i])
                self.progress_var.set((i - start + 2) / (end - start + 1) * 100)
                self.root.update_idletasks()
            
            
            self.status_var.set("Saving file...")
            with open(self.output_path.get(), 'wb') as output_file:
                writer.write(output_file)
            
            self.status_var.set("Extraction completed successfully!")
            messagebox.showinfo("Success", "PDF extraction completed successfully!")
            
        except Exception as e:
            self.status_var.set("Error occurred during extraction")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
        finally:
            self.progress_var.set(0)
            self.extract_button.configure(state='normal')

    def start_extraction(self):
        if not self.validate_inputs():
            return
            
        self.extract_button.configure(state='disabled')
        self.status_var.set("Starting extraction...")
        Thread(target=self.extract_pages, daemon=True).start()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def create_about_section(self):
        about_frame = ttk.Frame(self.main_frame)
        about_frame.pack(fill=tk.X, pady=(20, 0))
        
        contact_label = ttk.Label(
            about_frame,
            text="Contact: ",
            style='Modern.TLabel'
        )
        contact_label.pack(side=tk.LEFT)
        
        email_label = ttk.Label(
            about_frame,
            text="rijfas01@gmail.com",
            style='Link.TLabel',
            cursor="hand2"
        )
        email_label.pack(side=tk.LEFT)
        email_label.bind("<Button-1>", lambda e: webbrowser.open("mailto:rijfas01@gmail.com"))

    def show_about(self):
        about_text = """PDF Splitter v1.0
        
A modern GUI application for extracting pages from PDF files.

Created by: Rijfas
Contact: rijfas01@gmail.com
License: MIT License

© 2024 Rijfas. All rights reserved."""
        
        messagebox.showinfo("About PDF Splitter", about_text)

def main():
    root = ThemedTk(theme="arc")
    app = PDFSplitterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()