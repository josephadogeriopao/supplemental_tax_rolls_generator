"""
DOCX Supplemental Tax Rolls Generator - Main application runner
------------------
Description: Main execution script that defines input/output configurations
             and calls the external xlsx engine to process supplemental tax rolls data.

Author: Joseph Adogeri
Version: 1.0.0
Since: 2026-08-03
File: main.py
License: MIT
"""
import pandas as pd
import warnings
from dotenv import load_dotenv
import os
import sys

load_dotenv()

real_file = os.environ.get("REAL_XLSX_FILE", r"PATH TO PDF FILE")
pp_file = os.environ.get("PP_XLSX_FILE", r"PATH TO OUTPUT FOLDER/ DIRECTORY")

def main() -> None:
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    # Defensive check: Ensure both environment variables exist before running
    if not all([real_file, pp_file]):
        print("❌ Error: Missing configuration paths in your .env file or fallbacks!")
        return


    # Load the Excel files
    real_df = pd.read_excel(real_file)
    pp_df = pd.read_excel(pp_file)

    # 1. Force Pandas to show ALL rows and columns in the terminal
    pd.set_option('display.max_rows', None)  # None means unlimited rows
    pd.set_option('display.max_columns', None)  # None means unlimited columns
    pd.set_option('display.width', 1000)  # Prevents rows from wrapping into ugly lines

    # 2. Print the entire DataFrames (remove .head() to see more than 5 rows)
    print("--- REAL PROPERTY TAX ROLLS ---")
    print(real_df.head())
    print()
    print("--- PP PROPERTY TAX ROLLS ---")
    print(pp_df.head())

if __name__ == "__main__":
    main()
    sys.exit(0)
