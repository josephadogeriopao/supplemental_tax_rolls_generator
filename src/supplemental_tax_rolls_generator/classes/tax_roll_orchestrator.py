"""
DOCX Supplemental Tax Rolls Generator - Class-Based Architecture
-----------------------------------------------------------------
Description: High-level orchestration facade class that coordinates the structural
             lifecycle, workflow sequencing, and error checking of the pipeline execution.

Author: Joseph Adogeri
Version: 5.0.0
Since: 2026-08-21
File: tax_roll_orchestrator.py
License: MIT
"""

import sys
from .config_manager import ConfigManager
from .excel_data_appender import ExcelDataAppender
from .name_manager_updater import NameManagerUpdater


class TaxRollOrchestrator:
    """
    High-level facade class that coordinates the lifecycle of the engine pipeline stages.
    It links configuration context setups with raw data appending and name manager adjustments.
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
        Initializes the top-level orchestrator engine with specific timeline scopes.

        Args:
            target_year (int): The 4-digit processing base target tax year. Defaults to 2026.
            quarter (int): The target reporting calendar quarter (1-4). Defaults to 2.
        """
        self.config = ConfigManager(
            target_year=target_year,
            quarter=quarter,
            real_file=real_file,
            pp_file=pp_file,
            output_dir=output_dir
        )
    def execute_pipeline(self) -> None:
        """
        Runs validation, data cloning, tab insertions, and cell property overrides cleanly.

        This method guides the application step-by-step through checking configurations,
        duplicating template schemas, writing detailed sub-sheet database views, and forcing
        the final static variable override updates.

        Raises:
            SystemExit: Aborts script processing execution with exit status code 1 if file
                        locks or file system path exceptions are caught during validation passes.
        """
        if not self.config.validate():
            print("❌ Pipeline execution aborted due to initialization context failures.")
            sys.exit(1)

        # Step 1: Clone template
        ExcelDataAppender.clone_template(self.config.template_path, self.config.final_output_path)

        # Step 2: Append tables
        real_subs, pp_subs, real_col = ExcelDataAppender.append_all_data(
            self.config.final_output_path, self.config.real_file, self.config.pp_file
        )

        # Step 3: Run Named Ranges updates
        updater = NameManagerUpdater(self.config.final_output_path)
        updater.update_variables(self.config, real_subs, pp_subs, real_col)

        print("\n✅ Success! All pipeline segments executed and static values overwritten.")
