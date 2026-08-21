"""
DOCX Supplemental Tax Rolls Generator - Main application runner
----------------------------------------------------------------
Description: Main execution script that defines input/output configurations,
             writes structured details tabs, and builds the initial Name Manager
             variable mapping map using 'get_tax_roll_initial_values'. It safely
             overwrites worksheet-scoped local variables directly with raw computed
             numeric constants instead of cross-sheet dynamic formulas.

Author: Joseph Adogeri
Version: 4.8.0
Since: 2026-08-03
File: named_manager.py
License: MIT
"""


def get_tax_roll_initial_values(
    quarter: int, tax_year: int, str_date: str
) -> dict:
    """
    Generates the update configuration dictionary for worksheet-scoped variables.

    This function compiles a baseline mapping of core tax configuration attributes,
    homestead exemption buckets, and real estate categories used to update local
    Name Manager entries on the CONSOLIDATED SUMMARY worksheet.

    Args:
        quarter (int): The target reporting period (e.g., 1, 2, 3, or 4).
        tax_year (int): The 4-digit numeric baseline tax year (e.g., 2025, 2026).
        str_date (str): Uppercase, cross-platform formatted date string
                        (e.g., 'AUGUST 21, 2026').

    Returns:
        dict: A lookup map matching Name Manager variable keys to their initial
              numeric constants or string values.
    """
    return {
        "HOMESTEAD_EXEMPTION_NET_1": 0,
        "HOMESTEAD_EXEMPTION_NET_2": 0,
        "HOMESTEAD_EXEMPTION_NET_3": 0,
        "HOMESTEAD_EXEMPTION_NET_4": 0,
        "PERSONAL_PROPERTY_1": 0,
        "PERSONAL_PROPERTY_2": 0,
        "PERSONAL_PROPERTY_3": 0,
        "PERSONAL_PROPERTY_4": 0,
        "QUARTER": quarter,
        "REAL_ESTATE_1": 0,
        "REAL_ESTATE_2": 0,
        "REAL_ESTATE_3": 0,
        "REAL_ESTATE_4": 0,
        "TAX_YEAR": tax_year,
        "STR_DATE": str_date,
    }
