"""
DOCX Supplemental Tax Rolls Generator - Main application runner
-----------------------------------------------------------------
Description: Main execution entry script that initializes the object-oriented
             TaxRollOrchestrator pipeline facade. It initiates validation,
             re-cloning operations, and local Name Manager injections using
             the target tax year and reporting quarter scopes.

Author: Joseph Adogeri
Version: 5.0.0
Since: 2026-08-03
File: main.py
License: MIT
"""

import sys
import os
import warnings
from dotenv import load_dotenv

# Import the engine module from your separate local processing file
from classes.tax_roll_orchestrator import TaxRollOrchestrator

load_dotenv()


def main() -> None:
    """
    Main execution runtime function.

    Suppresses framework warnings, instantiates the master orchestrator class with
    configured baseline execution constants, and triggers the multi-layered data extraction
    and cell property update pipeline.
    """
    warnings.filterwarnings("ignore")

    # 🌍 Step 1: Explicitly load environments at the application root boundaries
    load_dotenv()

    # 🌍 Step 2: Resolve runtime targets and configurations
    TARGET_TAX_YEAR = 2026
    QUARTER = 2

    env_real_file = os.environ.get("REAL_XLSX_FILE")
    env_pp_file = os.environ.get("PP_XLSX_FILE")
    env_output_dir = os.environ.get("OUTPUT_DIR")

    # 🌍 Step 3: Forward resolved configurations straight to the orchestration layer
    orchestrator = TaxRollOrchestrator(
        target_year=TARGET_TAX_YEAR,
        quarter=QUARTER,
        real_file=env_real_file,
        pp_file=env_pp_file,
        output_dir=env_output_dir
    )

    orchestrator.execute_pipeline()


if __name__ == "__main__":
    main()
    sys.exit(0)
