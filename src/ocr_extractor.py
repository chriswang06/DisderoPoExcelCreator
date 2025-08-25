"""
OCR Extraction Module
Handles text extraction and parsing from images
"""

import re
from typing import Dict, List, Optional
from PIL import Image
import pytesseract


class OCRExtractor:
    def extract_text(self, image: Image.Image) -> str:
        """
        Extract text from image using OCR

        Args:
            image: PIL Image object

        Returns:
            Extracted text string
        """
        return pytesseract.image_to_string(image)

    def extract_po_number(self, text: str) -> Optional[str]:
        """
        Extract PO number from text

        Args:
            text: OCR text

        Returns:
            PO number or None if not found
        """
        po_pattern = r'LUMBER CO\.?\s+D(\d+)'
        match = re.search(po_pattern, text)
        if match:
            po_number = match.group(1).lstrip('0')
            return f'D{po_number}'
        return None

    def extract_product_blocks(self, text: str) -> List[str]:
        """
        Extract product blocks from OCR text

        Args:
            text: OCR text

        Returns:
            List of product block strings
        """
        lines = text.split('\n')
        blocks = []
        current_block = []
        in_product_block = False

        for line in lines:
            # Check if this is the start of a product block
            if re.match(r'^\d+\s+L[FE]\s+\d{6}-\d{4}-[A-Z]+', line):
                if in_product_block and current_block:
                    blocks.append('\n'.join(current_block))
                current_block = [line]
                in_product_block = True
            elif in_product_block:
                current_block.append(line)
                # Check if this is the end of a product block
                if re.search(r'\d+/\d+\'', line):
                    blocks.append('\n'.join(current_block))
                    in_product_block = False
                    current_block = []

        # Add any remaining block
        if in_product_block and current_block:
            blocks.append('\n'.join(current_block))

        return blocks

    def parse_product_block(self, block: str) -> Dict[str, Optional[str]]:
        """
        Parse a single product block

        Args:
            block: Product block text

        Returns:
            Dictionary with product information
        """
        result = {
            'product_code': None,
            'dimensions': None,
            'size': None
        }

        # Extract product code
        product_code_pattern = r'\b(\d{6}-\d{4}-[A-Z]+)\b'
        code_match = re.search(product_code_pattern, block)
        if code_match:
            result['product_code'] = code_match.group(1)

        # Extract dimensions
        dimensions_pattern = r'(\d+/\d+\'(?:,\s*\d+/\d+\')*)'
        dimensions_match = re.search(dimensions_pattern, block)
        if dimensions_match:
            result['dimensions'] = dimensions_match.group(1)

        # Extract size
        size_pattern = r'^\s*([\d.]+\s*[Xx]\s*[\d.]+)'
        size_match = re.search(size_pattern, block, re.MULTILINE)
        if size_match:
            result['size'] = size_match.group(1)

        return result

    def parse_document(self, text: str) -> Dict[str, List[Dict[str, Optional[str]]]]:
        """
        Parse entire document

        Args:
            text: OCR text

        Returns:
            Dictionary with PO number as key and list of products as value
        """
        po_number = self.extract_po_number(text)
        if not po_number:
            po_number = 'UNKNOWN'

        blocks = self.extract_product_blocks(text)
        products = []

        for block in blocks:
            parsed = self.parse_product_block(block)
            if parsed['product_code']:
                # Expand products with multiple dimensions
                if parsed['dimensions']:
                    dimensions = [dim.strip() for dim in parsed['dimensions'].split(',')]
                    for dimension in dimensions:
                        products.append({
                            'product_code': parsed['product_code'],
                            'dimensions': dimension,
                            'size': parsed['size']
                        })
                else:
                    products.append(parsed)

        return {po_number: products}