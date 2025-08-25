"""
Product Matching Module
Matches extracted products with master product list
"""

import re
from typing import List, Dict, Any, Optional
import pandas as pd


class ProductMatcher:
    def __init__(self, master_file: str):
        """
        Initialize product matcher with master product list

        Args:
            master_file: Path to master product Excel file
        """
        self.master_df = self._load_master_list(master_file)

    def _load_master_list(self, file_path: str) -> pd.DataFrame:
        """
        Load and process master product list

        Args:
            file_path: Path to Excel file

        Returns:
            Processed DataFrame
        """
        df = pd.read_excel(file_path, sheet_name='final')

        # Extract product code from description
        df['Product_Code'] = df['PRODUCT DESCRIPTION'].apply(self._extract_product_code)

        # Extract dimension length
        df['Dimension_Length'] = df['Dimension'].apply(self._extract_dimension_length)

        return df

    def _extract_product_code(self, description: Any) -> Optional[str]:
        """
        Extract product code from description

        Args:
            description: Product description string

        Returns:
            Product code or None
        """
        if pd.isna(description):
            return None

        # Pattern: 6 digits - 4 digits - letter(s)
        match = re.search(r'^(\d{6}-\d{4}-[A-Z]+)', str(description))
        if match:
            return match.group(1)

        # Alternative: split by newline and check first part
        parts = str(description).split('\n')
        if parts and re.match(r'\d{6}-\d{4}-[A-Z]+', parts[0]):
            return parts[0]

        return None

    def _extract_dimension_length(self, dimension: Any) -> Optional[str]:
        """
        Extract dimension length from dimension string

        Args:
            dimension: Dimension string

        Returns:
            Dimension length or None
        """
        if pd.isna(dimension):
            return None

        dim_str = str(dimension)

        # Remove quotes if present
        dim_str = dim_str.replace("'", "").replace('"', '')

        # If it contains *, split and get the last part
        if '*' in dim_str:
            parts = dim_str.split('*')
            return parts[-1].strip()
        else:
            # If no *, it's already just the length
            return dim_str.strip()

    def match_products(self, products: List[Dict[str, str]]) -> List[Dict[str, Any]]:
        """
        Match extracted products with master list

        Args:
            products: List of extracted products

        Returns:
            List of matched products with full information
        """
        # Convert to DataFrame for easier processing
        products_df = pd.DataFrame(products)

        if products_df.empty:
            return []

        # Split dimensions into piece count and length
        products_df[["Piece_Count", "Dimension_Length"]] = products_df["dimensions"].str.split("/", expand=True)
        products_df["Piece_Count"] = products_df["Piece_Count"].astype(int)

        # Clean dimension length
        products_df["Dimension_Length"] = (
            products_df["Dimension_Length"]
            .str.replace(r"[^\d]", "", regex=True)
            .astype(int)
        )

        # Clean master dimension length too
        self.master_df["Dimension_Length"] = (
            self.master_df["Dimension_Length"]
            .astype(str)
            .str.replace(r"[^\d]", "", regex=True)
        )
        # Convert to int, handling empty strings
        self.master_df["Dimension_Length"] = pd.to_numeric(
            self.master_df["Dimension_Length"], errors='coerce'
        ).fillna(0).astype(int)

        # Merge with master DataFrame
        merged = products_df.merge(
            self.master_df,
            left_on=["product_code", "Dimension_Length"],
            right_on=["Product_Code", "Dimension_Length"],
            how="left"
        )

        # Format quantity
        merged["Final_Quantity"] = merged.apply(self._format_quantity, axis=1)

        # Create final product list
        final_products = merged.apply(
            lambda row: {
                "product_code": row["product_code"],
                "SKU#": row["SKU#"] if pd.notna(row["SKU#"]) else "",
                "Product_Description": row["PRODUCT DESCRIPTION"] if pd.notna(row["PRODUCT DESCRIPTION"]) else "",
                "Dimension_Length": row["Dimension_Length"],
                "Quantity": row["Final_Quantity"] if pd.notna(row["Final_Quantity"]) else "",
                "Size": row["size"] if pd.notna(row["size"]) else ""
            },
            axis=1
        ).tolist()

        return final_products

    def _format_quantity(self, row: pd.Series) -> str:
        """
        Format quantity string based on piece count and unit quantity

        Args:
            row: DataFrame row

        Returns:
            Formatted quantity string
        """
        if pd.isna(row["QUANTITY"]):
            return ""

        match = re.search(r"(\d+)PC", str(row["QUANTITY"]))
        if match:
            per_unit = int(match.group(1))
            units = row["Piece_Count"] // per_unit
            return f"{units} {row['QUANTITY']}"

        return str(row["QUANTITY"])