"""
DOCX Supplemental Tax Rolls Generator - Class-Based Architecture
-----------------------------------------------------------------
Description: Class for data appending, managing template workbook replication,
             and inserting live Pandas data frames into distinct detail layers.
             Safely handles cases where either input stream is missing or empty.

Author: Joseph Adogeri
Version: 5.1.0
Since: 2026-08-21
File: excel_data_appender.py
License: MIT
"""

import os
import sys
import shutil
import warnings
from typing import List, Dict, Tuple, Any
import pandas as pd
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from dotenv import load_dotenv

# Pipeline engine imports
from excel_engine import append_data_and_details
from utils.date import format_date
from utils.named_manager import get_tax_roll_initial_values

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class ExcelDataAppender:
    """
    Manages cloning template workbooks and appending live Pandas data table sheets.
    Provides utility operations to bridge raw source data frames with target Excel sheets.
    """

    @staticmethod
    def clone_template(template_path: str, output_path: str) -> None:
        """
        Replicates the baseline master Excel template to the final target output path.py destination.

        Args:
            template_path (str): File system path.py to the master spreadsheet template.
            output_path (str): Destination file system path.py where the report is generated.
        """
        print("--- STEP 1: CLONING BASE TEMPLATE WORKBOOK STRUCTURE ---")
        shutil.copyfile(template_path, output_path)

    @staticmethod
    def append_all_data(output_path: str, real_file: str, pp_file: str) -> Tuple[Dict, Dict, str]:
        """
        Reads raw spreadsheet source streams and appends distinct data tables and SQL logs into the workbook.

        This method verifies file availability dynamically to handle partial pipeline execution runs,
        resolves structural evaluation column discrepancies, and updates the file in append-replace mode.

        Args:
            output_path (str): File system path.py to the cloned destination spreadsheet.
            real_file (str): File path.py containing incoming raw Real Property datasets.
            pp_file (str): File path.py containing incoming raw Personal Property datasets.

        Returns:
            Tuple[Dict, Dict, str]: A tuple containing:
                - Dict: Parsed subtotal matrix maps for Real Property records.
                - Dict: Parsed subtotal matrix maps for Personal Property records.
                - str: Resolved string column identifier used for tracking evaluation differences.
        """
        print("\n--- STEP 2: APPENDING LIVE DATA TABLES AND DETAIL SHEETS ---")

        # Initialize safe default variables
        real_subtotals = {}
        pp_subtotals = {}
        real_diff_col = "TOTAL_ASMT_DIFF"
        real_df = None
        pp_df = None

        # 🏢 Load Real Property dataset dynamically if a path.py is provided
        if real_file:
            print("Reading Real Property dataset...")
            real_df = pd.read_excel(real_file, sheet_name=0)
            real_diff_col = "TOTAL_ASMT_DIFF" if "TOTAL_ASMT_DIFF" in real_df.columns else "ASMT_TOTAL_DIFF"
        else:
            print("ℹ️ Skipping Real Property: No source file provided.")

        # 📦 Load Personal Property dataset dynamically if a path.py is provided
        if pp_file:
            print("Reading Personal Property dataset...")
            pp_df = pd.read_excel(pp_file, sheet_name=0)
        else:
            print("ℹ️ Skipping Personal Property: No source file provided.")

        # Execute structural sheet updates using the active openpyxl writer stream
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            if real_df is not None:
                print("Processing Real Property data...")
                real_subtotals = append_data_and_details(
                    writer, real_df, "REAL PROPERTY DETAILS", "REAL PROPERTY SQL", real_file
                )
                print("Processing Real Data...", real_subtotals)

            if pp_df is not None:
                print("Processing PP Property data...")
                pp_subtotals = append_data_and_details(
                    writer, pp_df, "PERSONAL PROPERTY DETAILS", "PERSONAL PROPERTY SQL", pp_file
                )
                print("Processing PP Data...", pp_subtotals)

        return real_subtotals, pp_subtotals, real_diff_col
