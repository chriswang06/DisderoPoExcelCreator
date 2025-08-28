"""
runtime_config.py
Runtime configuration for finding Tesseract and Poppler
Place this in your src/ folder
"""

import os
import sys
from pathlib import Path


def find_tesseract():
    """Find Tesseract executable in common installation locations"""
    tesseract_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'C:\Users\{}\AppData\Local\Tesseract-OCR\tesseract.exe'.format(os.environ.get('USERNAME', '')),
    ]

    # Check environment variable first
    env_path = os.environ.get('TESSERACT_PATH')
    if env_path and os.path.exists(env_path):
        return env_path

    # Check common paths
    for path in tesseract_paths:
        if os.path.exists(path):
            return path

    # Check if tesseract is in PATH
    import shutil
    tesseract_cmd = shutil.which('tesseract')
    if tesseract_cmd:
        return tesseract_cmd

    return None


def find_poppler():
    """Find Poppler bin directory in common installation locations"""
    poppler_paths = [
        r'C:\Program Files\poppler-25.07.0\Library\bin',
        r'C:\Program Files\poppler-23.08.0\Library\bin',
        r'C:\Program Files (x86)\poppler-25.07.0\Library\bin',
        r'C:\Program Files (x86)\poppler-23.08.0\Library\bin',
        r'C:\Program Files\poppler\Library\bin',
        r'C:\Program Files (x86)\poppler\Library\bin',
    ]

    # Check environment variable first
    env_path = os.environ.get('POPPLER_PATH')
    if env_path and os.path.exists(env_path):
        return env_path

    # Check common paths
    for path in poppler_paths:
        if os.path.exists(path) and os.path.exists(os.path.join(path, 'pdftoppm.exe')):
            return path

    # Check parent directories for any poppler installation
    program_files = [r'C:\Program Files', r'C:\Program Files (x86)']
    for pf in program_files:
        if os.path.exists(pf):
            for item in os.listdir(pf):
                if 'poppler' in item.lower():
                    potential_path = os.path.join(pf, item, 'Library', 'bin')
                    if os.path.exists(potential_path) and os.path.exists(os.path.join(potential_path, 'pdftoppm.exe')):
                        return potential_path

    return None


def configure_tools():
    """Configure Tesseract and Poppler paths"""
    import pytesseract

    # Configure Tesseract
    tesseract_path = find_tesseract()
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        print(f"Tesseract found at: {tesseract_path}")
    else:
        raise Exception(
            "Tesseract OCR not found!\n"
            "Please ensure Tesseract is installed.\n"
            "If installed in a custom location, set the TESSERACT_PATH environment variable."
        )

    # Configure and return Poppler path
    poppler_path = find_poppler()
    if poppler_path:
        print(f"Poppler found at: {poppler_path}")
    else:
        raise Exception(
            "Poppler not found!\n"
            "Please ensure Poppler is installed.\n"
            "If installed in a custom location, set the POPPLER_PATH environment variable."
        )

    return tesseract_path, poppler_path