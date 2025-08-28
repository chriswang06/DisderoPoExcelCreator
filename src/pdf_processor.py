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
    def __init__(self, dpi: int = 300, poppler_path=None):
        """
        Initialize PDF processor

        Args:
            dpi: Resolution for PDF to image conversion
            poppler_path: Path to Poppler binaries (for Windows)
        """
        self.dpi = dpi
        self.poppler_path = poppler_path
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

        # Convert PDF to images with poppler_path if provided
        try:
            if self.poppler_path:
                images = convert_from_path(
                    pdf_path,
                    dpi=self.dpi,
                    poppler_path=self.poppler_path
                )
            else:
                images = convert_from_path(pdf_path, dpi=self.dpi)
        except Exception as e:
            # Clean up temp directory if conversion fails
            if self.temp_dir:
                self.cleanup_temp_files()
            raise Exception(f"Failed to convert PDF to images: {str(e)}")

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
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                # Sometimes Windows holds onto files, just log the error
                print(f"Warning: Could not fully clean up temp files: {e}")
            finally:
                self.temp_dir = None
                self.temp_files = []