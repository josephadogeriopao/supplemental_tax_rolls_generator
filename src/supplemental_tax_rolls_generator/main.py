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

    # Point-of-entry pipeline initiation targeting specific parameters
    orchestrator = TaxRollOrchestrator(target_year=2026, quarter=2)
    orchestrator.execute_pipeline()


if __name__ == "__main__":
    main()
    sys.exit(0)
