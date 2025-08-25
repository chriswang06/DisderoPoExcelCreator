"""
Excel Report Generation Module
Generates formatted Excel reports for purchase orders
"""

import re
from datetime import date
from typing import List, Dict, Any
import pandas as pd
import xlsxwriter


class ExcelGenerator:
    def generate_report(self, po_number: str, products: List[Dict[str, Any]], output_file: str):
        """
        Generate Excel report for purchase order

        Args:
            po_number: Purchase order number
            products: List of matched products
            output_file: Output file path
        """
        # Create header rows
        header_rows = self._create_header_rows(po_number)

        # Create Disdero number row
        disdero_row = [po_number, "", "", "", ""]

        # Convert products to rows
        product_rows = self._create_product_rows(products)

        # Create footer row
        footer_row = self._create_footer_row(products)

        # Combine all rows
        all_rows = header_rows + [disdero_row] + product_rows + [footer_row]

        # Create DataFrame
        df = pd.DataFrame(all_rows)

        # Write to Excel with formatting
        self._write_formatted_excel(df, output_file, po_number)

    def _create_header_rows(self, po_number: str) -> List[List[str]]:
        """
        Create header rows for the report

        Args:
            po_number: Purchase order number

        Returns:
            List of header rows
        """
        return [
            [f"Load release to Disdero for PO#{po_number}", "", "", "", ""],
            ["Release to: DLC-2", "DISDERO LUMBER COMPANY", "", "", ""],
            ["", "12301 SE CARPENTER DRIVE", "", "", ""],
            ["", "CLACKAMAS, OR 97015", "", "", ""],
            ["CONTACT:", "503-239-8888 COURTNEY WARDELL", "", "", ""],
            ["", "", "", "", f"Date: {date.today()}"],
            ["Disdero #", "Dimension", "SKU#", "PRODUCT DESCRIPTION", "QUANTITY"],
        ]

    def _create_product_rows(self, products: List[Dict[str, Any]]) -> List[List[str]]:
        """
        Create product rows for the report

        Args:
            products: List of products

        Returns:
            List of product rows
        """
        rows = []
        for p in products:
            # Normalize size by removing spaces and converting X to *
            if p['Size']:
                # Remove any spaces around X/x and convert to uppercase
                normalized_size = re.sub(r'\s*[Xx]\s*', 'X', p['Size'])
                dimension = f"{normalized_size.replace('X', '*')}*{p['Dimension_Length']}"
            else:
                dimension = str(p['Dimension_Length'])

            rows.append([
                "",
                dimension,
                p["SKU#"],
                p["Product_Description"],
                p["Quantity"]
            ])
        return rows

    def _create_footer_row(self, products: List[Dict[str, Any]]) -> List[str]:
        """
        Create footer row with total units

        Args:
            products: List of products

        Returns:
            Footer row
        """
        total_units = sum(self._extract_units(p["Quantity"]) for p in products)
        return ["", "", "", f"** {total_units} UNITS TOTAL **", ""]

    def _extract_units(self, qty_str: str) -> int:
        """
        Extract unit count from quantity string

        Args:
            qty_str: Quantity string

        Returns:
            Unit count
        """
        parts = str(qty_str).split()
        return int(parts[0]) if parts and parts[0].isdigit() else 0

    def _write_formatted_excel(self, df: pd.DataFrame, output_file: str, po_number: str):
        """
        Write DataFrame to Excel with formatting

        Args:
            df: DataFrame to write
            output_file: Output file path
            po_number: Purchase order number
        """
        num_rows = len(df)

        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)

            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]

            # Set row heights
            for i in range(num_rows):
                worksheet.set_row(i, 27)

            # Define formats
            title_format = workbook.add_format({
                "bold": True,
                "font_size": 18,
                "align": "center",
                "valign": "vcenter"
            })

            bold_format = workbook.add_format({
                "bold": True,
                "font_size": 12
            })

            sub_header_format = workbook.add_format({
                "font_size": 12
            })

            header_format = workbook.add_format({
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#D9D9D9",
                "border": 1
            })

            table_format = workbook.add_format({
                'text_wrap': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            footer_format = workbook.add_format({
                'bold': True,
                'font_color': 'red',
                'bg_color': 'yellow',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            # Apply formatting
            # Title row
            worksheet.merge_range("A1:E1", df.iloc[0, 0], title_format)

            # Sub-header rows
            worksheet.write("A2", df.iloc[1, 0], bold_format)
            worksheet.write("B2", df.iloc[1, 1], sub_header_format)
            worksheet.write("B3", df.iloc[2, 1], sub_header_format)
            worksheet.write("B4", df.iloc[3, 1], sub_header_format)
            worksheet.write("A5", df.iloc[4, 0], bold_format)
            worksheet.write("B5", df.iloc[4, 1], sub_header_format)
            worksheet.write("E6", df.iloc[5, 4], bold_format)

            # Header row
            for col_num in range(5):
                worksheet.write(6, col_num, df.iloc[6, col_num], header_format)

            # Disdero number
            worksheet.write("A8", df.iloc[7, 0], bold_format)

            # Table data
            for col_num in range(5):
                for row_num in range(7, num_rows):
                    worksheet.write(row_num, col_num, df.iloc[row_num, col_num], table_format)

            # Footer
            worksheet.write(num_rows - 1, 3, df.iloc[num_rows - 1, 3], footer_format)

            # Set column widths
            worksheet.set_column("A:A", 10)
            worksheet.set_column("B:B", 16)
            worksheet.set_column("C:C", 22)
            worksheet.set_column("D:D", 47)
            worksheet.set_column("E:E", 25)