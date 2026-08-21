"""
DOCX Supplemental Tax Rolls Generator - Global Executable Entry Point
------------------
Description: Main application launcher. Restores user skin configurations
             and mounts decoupled layout components.
"""
import customtkinter as ctk
import os
import json
import sys

# Ensure your local src package modules are discoverable
sys.path.append(os.path.join(os.path.dirname(__file__), "src"))
from supplemental_tax_rolls_generator.ui.dashboard import TaxRollApp

CONFIG_FILE = "app_config.json"

def get_saved_appearance_mode():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f).get("appearance_mode", "System")
        except Exception:
            pass
    return "System"

def main():
    # Load appearance properties safely before building widgets to avoid screen flashing
    ctk.set_appearance_mode(get_saved_appearance_mode())
    ctk.set_default_color_theme("blue")

    app = TaxRollApp()
    app.mainloop()

if __name__ == "__main__":
    main()
