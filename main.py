"""
PDF to Word Converter GUI Application
A user-friendly interface for converting PDF files to Word documents
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from converter import PDFToWordConverter
import threading


class PDFToWordConverterGUI:
    """GUI Application for PDF to Word conversion"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Word Converter")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # Set color scheme
        self.bg_color = "#f0f0f0"
        self.accent_color = "#0078d4"
        self.success_color = "#107c10"
        self.error_color = "#d13438"
        
        self.root.configure(bg=self.bg_color)
        
        # Initialize converter
        self.converter = PDFToWordConverter()
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.is_converting = False
        
        # Setup GUI
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the GUI components"""
        # Title
        title_frame = tk.Frame(self.root, bg=self.accent_color)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        
        title_label = tk.Label(
            title_frame,
            text="PDF to Word Converter",
            font=("Arial", 20, "bold"),
            bg=self.accent_color,
            fg="white",
            pady=15
        )
        title_label.pack()
        
        # Main content frame
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # File selection section
        file_frame = tk.LabelFrame(main_frame, text="Select PDF File", font=("Arial", 11, "bold"), bg=self.bg_color)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        pdf_select_btn = tk.Button(
            file_frame,
            text="Browse PDF",
            command=self.select_pdf,
            bg=self.accent_color,
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2"
        )
        pdf_select_btn.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.pdf_label = tk.Label(
            file_frame,
            text="No file selected",
            font=("Arial", 9),
            bg=self.bg_color,
            fg="#666"
        )
        self.pdf_label.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.X, expand=True)
        
        # Output location section
        output_frame = tk.LabelFrame(main_frame, text="Output Location", font=("Arial", 11, "bold"), bg=self.bg_color)
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        output_default_btn = tk.Button(
            output_frame,
            text="Same Folder as PDF",
            command=lambda: self.output_path.set("same"),
            bg="#0a0a0a" if self.output_path.get() == "same" else "#e1e1e1",
            fg="white" if self.output_path.get() == "same" else "black",
            font=("Arial", 10),
            padx=15,
            pady=8,
            cursor="hand2"
        )
        output_default_btn.pack(side=tk.LEFT, padx=10, pady=10)
        
        output_custom_btn = tk.Button(
            output_frame,
            text="Choose Folder",
            command=self.select_output_folder,
            bg="#107c10" if self.output_path.get() != "same" else "#e1e1e1",
            fg="white" if self.output_path.get() != "same" else "black",
            font=("Arial", 10),
            padx=15,
            pady=8,
            cursor="hand2"
        )
        output_custom_btn.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.output_label = tk.Label(
            output_frame,
            text="Output: Same folder as PDF",
            font=("Arial", 9),
            bg=self.bg_color,
            fg="#666"
        )
        self.output_label.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.X, expand=True)
        
        # Set default output
        self.output_path.set("same")
        
        # Conversion options
        options_frame = tk.LabelFrame(main_frame, text="Conversion Options", font=("Arial", 11, "bold"), bg=self.bg_color)
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.keep_formatting = tk.BooleanVar(value=True)
        format_check = tk.Checkbutton(
            options_frame,
            text="Preserve PDF Formatting",
            variable=self.keep_formatting,
            font=("Arial", 10),
            bg=self.bg_color,
            cursor="hand2"
        )
        format_check.pack(anchor=tk.W, padx=10, pady=8)
        
        # Progress section
        progress_frame = tk.Frame(main_frame, bg=self.bg_color)
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='indeterminate',
            length=400
        )
        self.progress_bar.pack(fill=tk.X)
        
        self.status_label = tk.Label(
            progress_frame,
            text="Ready",
            font=("Arial", 9),
            bg=self.bg_color,
            fg="#666"
        )
        self.status_label.pack(pady=(10, 0))
        
        # Convert button
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.convert_btn = tk.Button(
            button_frame,
            text="Convert to Word",
            command=self.convert_pdf,
            bg=self.success_color,
            fg="white",
            font=("Arial", 12, "bold"),
            padx=40,
            pady=12,
            cursor="hand2"
        )
        self.convert_btn.pack(expand=True)
        
        # Output/Log section
        log_frame = tk.LabelFrame(main_frame, text="Conversion Log", font=("Arial", 11, "bold"), bg=self.bg_color)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 0))
        
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(
            log_frame,
            height=8,
            font=("Arial", 9),
            bg="white",
            fg="black",
            yscrollcommand=scrollbar.set
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        scrollbar.config(command=self.log_text.yview)
        
        self.log_text.insert(tk.END, "Welcome to PDF to Word Converter!\n")
        self.log_text.insert(tk.END, "Select a PDF file and click 'Convert to Word' to begin.\n")
        self.log_text.config(state=tk.DISABLED)
    
    def select_pdf(self):
        """Open file dialog to select PDF"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.pdf_path.set(file_path)
            file_name = os.path.basename(file_path)
            self.pdf_label.config(text=file_name, fg="black")
            self.log_message(f"Selected: {file_name}")
    
    def select_output_folder(self):
        """Open dialog to select output folder"""
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        
        if folder_path:
            self.output_path.set(folder_path)
            folder_name = os.path.basename(folder_path)
            self.output_label.config(text=f"Output: {folder_name}", fg="black")
            self.log_message(f"Output folder: {folder_name}")
    
    def convert_pdf(self):
        """Convert selected PDF to Word"""
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file first")
            return
        
        # Run conversion in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self._perform_conversion)
        thread.start()
    
    def _perform_conversion(self):
        """Perform the actual conversion"""
        try:
            self.is_converting = True
            self.convert_btn.config(state=tk.DISABLED)
            self.progress_bar.start()
            self.status_label.config(text="Converting...", fg="#0078d4")
            
            pdf_file = self.pdf_path.get()
            
            # Determine output path
            if self.output_path.get() == "same":
                output_file = None
            else:
                pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
                output_file = os.path.join(self.output_path.get(), f"{pdf_name}.docx")
            
            # Perform conversion
            success, message = self.converter.convert(pdf_file, output_file)
            
            self.progress_bar.stop()
            self.is_converting = False
            
            if success:
                self.status_label.config(text="Conversion successful!", fg=self.success_color)
                self.log_message("✓ Conversion successful!")
                self.log_message(message)
                messagebox.showinfo("Success", message)
            else:
                self.status_label.config(text="Conversion failed", fg=self.error_color)
                self.log_message("✗ Conversion failed!")
                self.log_message(message)
                messagebox.showerror("Error", message)
        
        except Exception as e:
            self.progress_bar.stop()
            self.is_converting = False
            self.status_label.config(text="Error occurred", fg=self.error_color)
            self.log_message(f"✗ Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        finally:
            self.convert_btn.config(state=tk.NORMAL)
    
    def log_message(self, message: str):
        """Add message to log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()


def main():
    """Main entry point"""
    root = tk.Tk()
    app = PDFToWordConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
