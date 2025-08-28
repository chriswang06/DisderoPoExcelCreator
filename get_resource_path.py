# get_resource_path.py
# Add this to your project and import it in gui_app.py

import sys
import os

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Example usage in your gui_app.py:
# from get_resource_path import get_resource_path
#
# products_file = get_resource_path('productslist.xlsx')
# src_path = get_resource_path('src')