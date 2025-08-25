"""
Configuration Module
Stores application configuration and constants
"""

from dataclasses import dataclass
from typing import Optional


@dataclass
class Config:
    """Application configuration"""

    # Company information
    COMPANY_NAME = "DISDERO LUMBER COMPANY"
    COMPANY_ADDRESS_LINE1 = "12301 SE CARPENTER DRIVE"
    COMPANY_ADDRESS_LINE2 = "CLACKAMAS, OR 97015"
    COMPANY_CONTACT = "503-239-8888"
    COMPANY_CONTACT_NAME = "COURTNEY WARDELL"
    RELEASE_TO = "DLC-2"

    # OCR settings
    DEFAULT_DPI = 300

    # Excel formatting colors
    HEADER_BG_COLOR = "#D9D9D9"
    FOOTER_BG_COLOR = "yellow"
    FOOTER_TEXT_COLOR = "red"

    # Column widths for Excel
    COLUMN_WIDTHS = {
        "A": 10,
        "B": 16,
        "C": 22,
        "D": 47,
        "E": 25
    }

    # Row height
    DEFAULT_ROW_HEIGHT = 27

    # Regex patterns
    PO_NUMBER_PATTERN = r'LUMBER CO\.?\s+D(\d+)'
    PRODUCT_CODE_PATTERN = r'\b(\d{6}-\d{4}-[A-Z]+)\b'
    DIMENSIONS_PATTERN = r'(\d+/\d+\'(?:,\s*\d+/\d+\')*)'
    SIZE_PATTERN = r'^\s*([\d.]+\s*[Xx]\s*[\d.]+)'  # Handles both "2X6" and "2 X 6"
    PRODUCT_BLOCK_START_PATTERN = r'^\d+\s+L[FE]\s+\d{6}-\d{4}-[A-Z]+'