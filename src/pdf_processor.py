"""
PDF Processing Module
Handles PDF to image conversion and image manipulation
"""

from pathlib import Path
from typing import List
from PIL import Image
from pdf2image import convert_from_path
import tempfile
import shutil


class PDFProcessor:
    def __init__(self, dpi: int = 300):
        """
        Initialize PDF processor

        Args:
            dpi: Resolution for PDF to image conversion
        """
        self.dpi = dpi
        self.temp_dir = None
        self.temp_files = []

    def convert_pdf_to_images(self, pdf_path: str) -> List[Image.Image]:
        """
        Convert PDF to list of PIL Image objects

        Args:
            pdf_path: Path to PDF file

        Returns:
            List of PIL Image objects, one per page
        """
        # Create temporary directory for intermediate files
        self.temp_dir = tempfile.mkdtemp(prefix="po_processor_")
        temp_path = Path(self.temp_dir)

        # Convert PDF to images
        images = convert_from_path(pdf_path, self.dpi)

        # Save images temporarily if needed for debugging
        saved_images = []
        for i, img in enumerate(images):
            temp_file = temp_path / f"page_{i}.jpg"
            img.save(temp_file, 'JPEG')
            self.temp_files.append(temp_file)
            saved_images.append(img)

        return saved_images

    def combine_images_vertically(self, images: List[Image.Image]) -> Image.Image:
        """
        Combine multiple images vertically into one

        Args:
            images: List of PIL Image objects

        Returns:
            Combined PIL Image object
        """
        if not images:
            raise ValueError("No images to combine")

        if len(images) == 1:
            return images[0]

        # Calculate total height and max width
        total_height = sum(img.height for img in images)
        max_width = max(img.width for img in images)

        # Create new image
        combined = Image.new('RGB', (max_width, total_height))

        # Paste images vertically
        y_offset = 0
        for img in images:
            combined.paste(img, (0, y_offset))
            y_offset += img.height

        return combined

    def cleanup_temp_files(self):
        """Clean up temporary files and directory"""
        if self.temp_dir and Path(self.temp_dir).exists():
            shutil.rmtree(self.temp_dir)
            self.temp_dir = None
            self.temp_files = []