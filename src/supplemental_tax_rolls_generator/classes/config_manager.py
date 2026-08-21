"""
DOCX Supplemental Tax Rolls Generator - Class-Based Architecture
-----------------------------------------------------------------
Description: Class managing system configuration environment initialization,
             filesystem path routing, and target timeline sequence builds.

Author: Joseph Adogeri
Version: 5.0.0
Since: 2026-08-21
File: config_manager.py
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


class ConfigManager:
    """
    Handles environment variable resolution, filesystem validation, and core constants
    required to run the supplemental tax rolls pipeline.
    """

    def __init__(
        self,
        target_year: int,
        quarter: int,
        real_file: Optional[str] = None,
        pp_file: Optional[str] = None,
        output_dir: Optional[str] = None
    ):
        """
        Initializes environmental configurations, loads dotenv files, and maps core parameters.

        Args:
            target_year (int): The active baseline processing tax year. Defaults to 2026.
            quarter (int): The target reporting period quarter (1-4). Defaults to 2.
        """
        self.target_tax_year = target_year
        self.quarter = quarter
        self.str_date = format_date()
        self.timeline_years = [self.target_tax_year - i for i in range(4)]

        # Load environment paths
        self.real_file = real_file if real_file else ""
        self.pp_file = pp_file if pp_file else ""
        self.output_dir = os.environ.get("OUTPUT_DIR", os.path.dirname(self.pp_file) if self.pp_file else "")

        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.template_path = os.path.join(self.script_dir, "..", "..", "..", "templates", "str_template.xlsx")
        self.final_output_path = os.path.join(self.output_dir, "consolidated supplemental tax rolls.xlsx")


        if self.output_dir:
            self.final_output_path = os.path.join(self.output_dir, "consolidated supplemental tax rolls.xlsx")
        else:
            self.final_output_path = ""

    def validate(self) -> bool:
        """
        Verifies configuration readiness, environment variables, template locations,
        and destination spreadsheet write accessibility.

        Returns:
            bool: True if all system configuration settings pass validation checks, False otherwise.
        """
        if not all([self.real_file, self.pp_file, self.output_dir]):
            print("❌ Error: Missing configuration paths or OUTPUT_DIR in your .env file!")
            return False

        if not os.path.exists(self.template_path):
            print(f"❌ Error: Base spreadsheet template not found at: {self.template_path}")
            return False

        os.makedirs(self.output_dir, exist_ok=True)
        return self._check_file_permission()

    def _check_file_permission(self) -> bool:
        """
        Ensures the destination path isn't locked by an open instance of Excel.

        Internal helper function that attempts a read/write hook on the destination file
        to capture OS PermissionError anomalies early.

        Returns:
            bool: True if the file can be modified or doesn't exist yet, False if locked.
        """
        if os.path.exists(self.final_output_path):
            try:
                with open(self.final_output_path, "r+"):
                    pass
            except PermissionError:
                print(f"\n❌ ERROR: Permission Denied to file: {self.final_output_path}")
                print("👉 Please close the output spreadsheet in Microsoft Excel and rerun the script.")
                return False
        return True
