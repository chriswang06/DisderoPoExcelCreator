#!/usr/bin/env python3
"""
Purchase Order Processor - GUI Application
A graphical interface for processing purchase order PDFs
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import threading
import sys
import traceback

# Import our modules
from src.pdf_processor import PDFProcessor
from src.ocr_extractor import OCRExtractor
from src.product_matcher import ProductMatcher
from src.excel_generator import ExcelGenerator


class POProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Purchase Order Processor")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        # Variables
        self.pdf_path = tk.StringVar()
        self.master_path = tk.StringVar(value="productslist.xlsx")
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        self.progress_text = tk.StringVar()

        # Create GUI elements
        self.create_widgets()

        # Set default output directory
        self.output_path.set(str(Path.home() / "Desktop"))

    def create_widgets(self):
        # Title
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="Purchase Order Processor",
            font=("Arial", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(expand=True)

        # Main content frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # Input PDF Selection
        input_frame = tk.LabelFrame(main_frame, text="Input PDF", padx=10, pady=10)
        input_frame.pack(fill="x", pady=(0, 10))

        tk.Entry(
            input_frame,
            textvariable=self.pdf_path,
            width=50,
            state="readonly"
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            input_frame,
            text="Browse",
            command=self.browse_pdf,
            width=10
        ).pack(side="left")

        # Master File Selection
        master_frame = tk.LabelFrame(main_frame, text="Master Product List", padx=10, pady=10)
        master_frame.pack(fill="x", pady=(0, 10))

        tk.Entry(
            master_frame,
            textvariable=self.master_path,
            width=50,
            state="readonly"
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            master_frame,
            text="Browse",
            command=self.browse_master,
            width=10
        ).pack(side="left")

        # Output Directory Selection
        output_frame = tk.LabelFrame(main_frame, text="Output Directory", padx=10, pady=10)
        output_frame.pack(fill="x", pady=(0, 10))

        tk.Entry(
            output_frame,
            textvariable=self.output_path,
            width=50,
            state="readonly"
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            output_frame,
            text="Browse",
            command=self.browse_output,
            width=10
        ).pack(side="left")

        # Options Frame
        options_frame = tk.LabelFrame(main_frame, text="Options", padx=10, pady=10)
        options_frame.pack(fill="x", pady=(0, 10))

        tk.Label(options_frame, text="DPI:").pack(side="left", padx=(0, 5))
        self.dpi_var = tk.IntVar(value=300)
        dpi_spinbox = tk.Spinbox(
            options_frame,
            from_=150,
            to=600,
            increment=50,
            textvariable=self.dpi_var,
            width=10
        )
        dpi_spinbox.pack(side="left", padx=(0, 20))

        tk.Label(options_frame, text="(Higher DPI = better quality but slower)").pack(side="left")

        # Process Button
        self.process_btn = tk.Button(
            main_frame,
            text="Process PDF",
            command=self.process_pdf,
            height=2,
            width=20,
            font=("Arial", 12, "bold"),
            bg="#27ae60",
            fg="white",
            cursor="hand2"
        )
        self.process_btn.pack(pady=20)

        # Progress Bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            mode="indeterminate",
            length=400
        )
        self.progress_bar.pack(pady=(0, 10))

        # Progress Text
        progress_label = tk.Label(
            main_frame,
            textvariable=self.progress_text,
            font=("Arial", 9),
            fg="#7f8c8d"
        )
        progress_label.pack()

        # Status Bar
        status_frame = tk.Frame(self.root, bg="#ecf0f1", height=30)
        status_frame.pack(fill="x", side="bottom")
        status_frame.pack_propagate(False)

        status_label = tk.Label(
            status_frame,
            textvariable=self.status_text,
            bg="#ecf0f1",
            anchor="w",
            padx=10
        )
        status_label.pack(fill="both", expand=True)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Select Purchase Order PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            self.status_text.set(f"Selected: {Path(filename).name}")

    def browse_master(self):
        filename = filedialog.askopenfilename(
            title="Select Master Product List",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.master_path.set(filename)
            self.status_text.set(f"Master file: {Path(filename).name}")

    def browse_output(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_path.set(directory)
            self.status_text.set(f"Output to: {directory}")

    def process_pdf(self):
        # Validate inputs
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return

        if not Path(self.pdf_path.get()).exists():
            messagebox.showerror("Error", "Selected PDF file does not exist")
            return

        if not Path(self.master_path.get()).exists():
            messagebox.showerror("Error", "Master product list file not found")
            return

        if not self.output_path.get():
            messagebox.showerror("Error", "Please select an output directory")
            return

        # Disable button during processing
        self.process_btn.config(state="disabled", bg="#95a5a6")

        # Start processing in a separate thread
        thread = threading.Thread(target=self.process_pdf_thread)
        thread.daemon = True
        thread.start()

    def process_pdf_thread(self):
        try:
            # Start progress bar
            self.progress_bar.start(10)

            # Step 1: Convert PDF to images
            self.update_progress("Converting PDF to images...")
            pdf_processor = PDFProcessor(dpi=self.dpi_var.get())
            images = pdf_processor.convert_pdf_to_images(self.pdf_path.get())

            # Step 2: Combine images and perform OCR
            self.update_progress("Performing OCR (this may take a moment)...")
            ocr_extractor = OCRExtractor()
            combined_image = pdf_processor.combine_images_vertically(images)
            ocr_text = ocr_extractor.extract_text(combined_image)

            # Step 3: Parse OCR results
            self.update_progress("Parsing document...")
            parsed_data = ocr_extractor.parse_document(ocr_text)

            # Step 4: Match products with master list
            self.update_progress("Matching products...")
            matcher = ProductMatcher(self.master_path.get())
            po_number, products = list(parsed_data.items())[0]
            matched_products = matcher.match_products(products)

            # Step 5: Generate Excel report
            self.update_progress(f"Generating Excel report for PO #{po_number}...")
            excel_gen = ExcelGenerator()
            output_file = Path(self.output_path.get()) / f"Disdero #{po_number}.xlsx"
            excel_gen.generate_report(po_number, matched_products, str(output_file))

            # Clean up
            pdf_processor.cleanup_temp_files()

            # Stop progress bar
            self.progress_bar.stop()

            # Show success message
            self.root.after(0, lambda: self.show_success(output_file, po_number))

        except Exception as e:
            # Stop progress bar
            self.progress_bar.stop()

            # Show error message
            error_msg = f"Error processing PDF:\n\n{str(e)}\n\nDetails:\n{traceback.format_exc()}"
            self.root.after(0, lambda: messagebox.showerror("Processing Error", error_msg))
            self.root.after(0, lambda: self.update_progress(""))
            self.root.after(0, lambda: self.status_text.set("Error occurred during processing"))

        finally:
            # Re-enable button
            self.root.after(0, lambda: self.process_btn.config(state="normal", bg="#27ae60"))

    def update_progress(self, message):
        self.progress_text.set(message)
        self.root.update_idletasks()

    def show_success(self, output_file, po_number):
        self.update_progress("")
        self.status_text.set(f"✓ Successfully processed PO #{po_number}")

        result = messagebox.askyesno(
            "Success",
            f"Report generated successfully!\n\nPO Number: {po_number}\nFile: {output_file.name}\n\nWould you like to open the output folder?"
        )

        if result:
            # Open the output folder in file explorer
            import subprocess
            import platform

            if platform.system() == 'Windows':
                subprocess.Popen(['explorer', str(output_file.parent)])
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', str(output_file.parent)])
            else:  # Linux
                subprocess.Popen(['xdg-open', str(output_file.parent)])


def main():
    root = tk.Tk()
    app = POProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()