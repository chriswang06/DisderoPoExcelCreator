#!/usr/bin/env python3
"""
Purchase Order Processing System
Main entry point for the application
"""

import argparse
import sys
from pathlib import Path
from src.pdf_processor import PDFProcessor
from src.ocr_extractor import OCRExtractor
from src.product_matcher import ProductMatcher
from src.excel_generator import ExcelGenerator
from src.config import Config


def main():
    parser = argparse.ArgumentParser(description='Process purchase order PDFs and generate Excel reports')
    parser.add_argument('pdf_path', help='Path to the purchase order PDF')
    parser.add_argument('--master-file', default='productslist.xlsx',
                        help='Path to master product list Excel file')
    parser.add_argument('--output-dir', default='output',
                        help='Directory for output files')
    parser.add_argument('--dpi', type=int, default=300,
                        help='DPI for PDF to image conversion')
    args = parser.parse_args()

    # Create output directory if it doesn't exist
    output_dir = Path(args.output_dir)
    output_dir.mkdir(exist_ok=True)

    try:
        # Step 1: Convert PDF to images
        print(f"Converting PDF: {args.pdf_path}")
        pdf_processor = PDFProcessor(dpi=args.dpi)
        images = pdf_processor.convert_pdf_to_images(args.pdf_path)

        # Step 2: Combine images and perform OCR
        print("Performing OCR...")
        ocr_extractor = OCRExtractor()
        combined_image = pdf_processor.combine_images_vertically(images)
        ocr_text = ocr_extractor.extract_text(combined_image)

        # Step 3: Parse OCR results
        print("Parsing document...")
        parsed_data = ocr_extractor.parse_document(ocr_text)

        # Step 4: Match products with master list
        print("Matching products...")
        matcher = ProductMatcher(args.master_file)
        po_number, products = list(parsed_data.items())[0]
        matched_products = matcher.match_products(products)

        # Step 5: Generate Excel report
        print(f"Generating Excel report for PO #{po_number}")
        excel_gen = ExcelGenerator()
        output_file = output_dir / f"Disdero #{po_number}.xlsx"
        excel_gen.generate_report(po_number, matched_products, str(output_file))

        print(f"✓ Report generated successfully: {output_file}")

        # Clean up temporary image files
        pdf_processor.cleanup_temp_files()

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()