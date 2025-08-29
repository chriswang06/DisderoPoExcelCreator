#!/usr/bin/env python3
"""
Test script for all src module functions
Run this to test the complete pipeline with your actual PDF
"""

import sys
import os
from pathlib import Path
from datetime import date
import pandas as pd

# Add src to path
sys.path.insert(0, 'src')

# Import all modules
from pdf_processor import PDFProcessor
from ocr_extractor import OCRExtractor
from product_matcher import ProductMatcher
from excel_generator import ExcelGenerator
from config import Config


def test_pdf_processor(pdf_path):
    """Test PDF processing functions"""
    print("=" * 60)
    print("TESTING PDF PROCESSOR")
    print("=" * 60)

    processor = PDFProcessor(dpi=300)

    # Test PDF to images conversion
    print("\n1. Converting PDF to images...")
    images = processor.convert_pdf_to_images(pdf_path)
    print(f"✓ Converted to {len(images)} images")
    for i, img in enumerate(images):
        print(f"  Page {i + 1}: {img.size[0]}x{img.size[1]} pixels")

    # Test image combination
    print("\n2. Combining images vertically...")
    combined = processor.combine_images_vertically(images)
    print(f"✓ Combined image size: {combined.size[0]}x{combined.size[1]} pixels")

    # Cleanup
    processor.cleanup_temp_files()
    print("✓ Temporary files cleaned up")

    return images, combined


def test_ocr_extractor(combined_image):
    """Test OCR extraction functions"""
    print("\n" + "=" * 60)
    print("TESTING OCR EXTRACTOR")
    print("=" * 60)

    extractor = OCRExtractor()

    # Test text extraction
    print("\n1. Extracting text from image...")
    ocr_text = extractor.extract_text(combined_image)
    print(f"✓ Extracted {len(ocr_text)} characters")

    # Save OCR text for inspection
    with open('test_ocr_output.txt', 'w') as f:
        f.write(ocr_text)
    print("✓ Saved OCR text to test_ocr_output.txt")

    # Test PO number extraction
    print("\n2. Extracting PO number...")
    po_number = extractor.extract_po_number(ocr_text)
    print(f"✓ PO Number: {po_number}")

    # Test product block extraction
    print("\n3. Extracting product blocks...")
    blocks = extractor.extract_product_blocks(ocr_text)
    print(f"✓ Found {len(blocks)} product blocks")

    # Test block parsing
    print("\n4. Parsing product blocks...")
    for i, block in enumerate(blocks[:3], 1):  # Show first 3
        parsed = extractor.parse_product_block(block)
        print(f"\n  Block {i}:")
        print(f"    Product Code: {parsed.get('product_code')}")
        print(f"    Size: {parsed.get('size')}")
        print(f"    Dimensions: {parsed.get('dimensions')}")

    # Test full document parsing
    print("\n5. Parsing complete document...")
    parsed_data = extractor.parse_document(ocr_text)
    po_number, products = list(parsed_data.items())[0]
    print(f"✓ PO Number: {po_number}")
    print(f"✓ Total products: {len(products)}")

    return ocr_text, po_number, products


def test_product_matcher(products, master_file='productslist.xlsx'):
    """Test product matching functions"""
    print("\n" + "=" * 60)
    print("TESTING PRODUCT MATCHER")
    print("=" * 60)

    if not os.path.exists(master_file):
        print(f"⚠ Master file not found: {master_file}")
        return []

    matcher = ProductMatcher(master_file)

    # Test master list loading
    print(f"\n1. Master list loaded:")
    print(f"✓ Total SKUs in master: {len(matcher.master_df)}")
    print(f"✓ Columns: {list(matcher.master_df.columns)}")

    # Test product matching
    print("\n2. Matching products...")
    matched_products = matcher.match_products(products)
    print(f"✓ Matched {len(matched_products)} products")

    # Show matching results
    print("\n3. Sample matching results:")
    for i, product in enumerate(matched_products[:5], 1):
        print(f"\n  Product {i}:")
        print(f"    Code: {product.get('product_code')}")
        print(f"    SKU: {product.get('SKU#')}")
        print(f"    Size: {product.get('Size')}")
        print(f"    Dimension: {product.get('Dimension_Length')}")
        print(f"    Quantity: {product.get('Quantity')}")

    # Check for unmatched products
    unmatched = [p for p in matched_products if not p.get('SKU#')]
    if unmatched:
        print(f"\n⚠ Unmatched products: {len(unmatched)}")
        for p in unmatched[:3]:
            print(f"  - {p.get('product_code')} / {p.get('Dimension_Length')}")

    return matched_products


def test_excel_generator(po_number, matched_products):
    """Test Excel report generation"""
    print("\n" + "=" * 60)
    print("TESTING EXCEL GENERATOR")
    print("=" * 60)

    generator = ExcelGenerator()

    # Generate test report
    output_file = f"test_Disdero #{po_number}.xlsx"
    print(f"\n1. Generating Excel report: {output_file}")

    generator.generate_report(po_number, matched_products, output_file)
    print(f"✓ Report generated successfully")

    # Verify file exists
    if os.path.exists(output_file):
        file_size = os.path.getsize(output_file) / 1024  # KB
        print(f"✓ File size: {file_size:.1f} KB")

        # Read back and verify
        df = pd.read_excel(output_file, header=None)
        print(f"✓ Report dimensions: {df.shape[0]} rows x {df.shape[1]} columns")

        # Check header
        print(f"✓ Title: {df.iloc[0, 0]}")
        print(f"✓ Date: {df.iloc[5, 4]}")

    return output_file


def test_config():
    """Test configuration values"""
    print("\n" + "=" * 60)
    print("TESTING CONFIG")
    print("=" * 60)

    print(f"\nCompany Info:")
    print(f"  Name: {Config.COMPANY_NAME}")
    print(f"  Address: {Config.COMPANY_ADDRESS_LINE1}")
    print(f"  Contact: {Config.COMPANY_CONTACT}")

    print(f"\nSettings:")
    print(f"  Default DPI: {Config.DEFAULT_DPI}")
    print(f"  Header Color: {Config.HEADER_BG_COLOR}")

    print(f"\nRegex Patterns:")
    print(f"  PO Pattern: {Config.PO_NUMBER_PATTERN}")
    print(f"  Product Code Pattern: {Config.PRODUCT_CODE_PATTERN}")


def run_complete_pipeline(pdf_path):
    """Run the complete processing pipeline"""
    print("\n" + "=" * 80)
    print("RUNNING COMPLETE PIPELINE")
    print("=" * 80)

    try:
        # Step 1: Process PDF
        print("\nStep 1: Processing PDF...")
        processor = PDFProcessor(dpi=300)
        images = processor.convert_pdf_to_images(pdf_path)
        combined = processor.combine_images_vertically(images)

        # Step 2: Extract text
        print("Step 2: Extracting text...")
        extractor = OCRExtractor()
        ocr_text = extractor.extract_text(combined)
        parsed_data = extractor.parse_document(ocr_text)
        po_number, products = list(parsed_data.items())[0]

        # Step 3: Match products
        print("Step 3: Matching products...")
        matcher = ProductMatcher('productslist.xlsx')
        matched_products = matcher.match_products(products)

        # Step 4: Generate report
        print("Step 4: Generating report...")
        generator = ExcelGenerator()
        output_file = f"Disdero #{po_number}.xlsx"
        generator.generate_report(po_number, matched_products, output_file)

        # Cleanup
        processor.cleanup_temp_files()

        print(f"\n✓ Pipeline complete!")
        print(f"✓ Report saved as: {output_file}")
        print(f"✓ Total products processed: {len(matched_products)}")

        # Calculate total units
        total_units = 0
        for p in matched_products:
            qty_str = str(p.get('Quantity', ''))
            if qty_str:
                parts = qty_str.split()
                if parts and parts[0].isdigit():
                    total_units += int(parts[0])
        print(f"✓ Total units: {total_units}")

        return output_file

    except Exception as e:
        print(f"\n✗ Pipeline failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Main test function"""
    # Get PDF path from command line or use default
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else 'test.pdf'

    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found: {pdf_path}")
        print("\nUsage: python test_src_functions.py [path/to/pdf]")
        return

    print(f"Testing with PDF: {pdf_path}\n")

    # Test individual modules
    images, combined = test_pdf_processor(pdf_path)
    ocr_text, po_number, products = test_ocr_extractor(combined)
    matched_products = test_product_matcher(products)

    if matched_products:
        test_excel_generator(po_number, matched_products)

    test_config()

    # Run complete pipeline
    print("\n" + "=" * 80)
    print("Now testing the complete pipeline...")
    output_file = run_complete_pipeline(pdf_path)

    if output_file:
        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
        print(f"✓ Final report: {output_file}")
        print("✓ OCR text saved: test_ocr_output.txt")
    else:
        print("\n✗ Some tests failed. Check the output above for details.")


if __name__ == "__main__":
    main()