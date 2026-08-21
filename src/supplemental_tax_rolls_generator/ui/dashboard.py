import customtkinter as ctk
import os
import threading
import time
import json
from tkinter import filedialog
from .dialogs import AnimatedDialog  # Import the new decoupled popup layer

CONFIG_FILE = "app_config.json"

class TaxRollApp(ctk.CTk):
    """Main workspace engine control dashboard constructed using CustomTkinter layouts."""

    def __init__(self):
        super().__init__()

        self.title("Supplemental Tax Rolls Generator Dashboard")
        self.geometry("700x580")
        self.resizable(False, False)

        self.real_file_path = ctk.StringVar(value="")
        self.pp_file_path = ctk.StringVar(value="")
        self.output_dir_path = ctk.StringVar(value="")

        self.grid_columnconfigure(0, weight=1)

        # 📌 Header Frame
        self.header_frame = ctk.CTkFrame(self, height=70, corner_radius=0, fg_color=("#3b5998", "#1e272c"))
        self.header_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))

        self.header_title = ctk.CTkLabel(
            self.header_frame, text="🏛️ Supplemental Tax Rolls",
            font=ctk.CTkFont(size=20, weight="bold"), text_color="white"
        )
        self.header_title.pack(side="left", padx=25, pady=15)

        self.theme_menu = ctk.CTkComboBox(
            self.header_frame, values=["System", "Light", "Dark"], width=100,
            command=self.change_appearance_theme
        )
        self.theme_menu.pack(side="right", padx=25, pady=15)

        # 📦 Central Fields Canvas Area
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=1, column=0, sticky="nsew", padx=30)

        self.lbl_real = ctk.CTkLabel(self.main_container, text="Real Property Spreadsheet (Optional):",
                                     font=ctk.CTkFont(weight="bold"))
        self.lbl_real.grid(row=0, column=0, sticky="w", pady=(10, 2))
        self.entry_real = ctk.CTkEntry(self.main_container, textvariable=self.real_file_path, width=480,
                                       placeholder_text="Select real property file location...")
        self.entry_real.grid(row=1, column=0, padx=(0, 10), pady=(0, 10))
        self.btn_real = ctk.CTkButton(self.main_container, text="Browse", width=120, command=self.browse_real_file)
        self.btn_real.grid(row=1, column=1, pady=(0, 10))

        self.lbl_pp = ctk.CTkLabel(self.main_container, text="Personal Property Spreadsheet (Optional):",
                                   font=ctk.CTkFont(weight="bold"))
        self.lbl_pp.grid(row=2, column=0, sticky="w", pady=(10, 2))
        self.entry_pp = ctk.CTkEntry(self.main_container, textvariable=self.pp_file_path, width=480,
                                     placeholder_text="Select personal property file location...")
        self.entry_pp.grid(row=3, column=0, padx=(0, 10), pady=(0, 10))
        self.btn_pp = ctk.CTkButton(self.main_container, text="Browse", width=120, command=self.browse_pp_file)
        self.btn_pp.grid(row=3, column=1, pady=(0, 10))

        self.lbl_out = ctk.CTkLabel(self.main_container, text="Target Output Folder (Required):",
                                    font=ctk.CTkFont(weight="bold"))
        self.lbl_out.grid(row=4, column=0, sticky="w", pady=(10, 2))
        self.entry_out = ctk.CTkEntry(self.main_container, textvariable=self.output_dir_path, width=480,
                                      placeholder_text="Select folder location where final report saves...")
        self.entry_out.grid(row=5, column=0, padx=(0, 10), pady=(0, 10))
        self.btn_out = ctk.CTkButton(self.main_container, text="Browse Target", width=120, fg_color="#2ecc71",
                                     hover_color="#27ae60", text_color="white", command=self.browse_output_dir)
        self.btn_out.grid(row=5, column=1, pady=(0, 10))

        self.param_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.param_frame.grid(row=6, column=0, columnspan=2, sticky="w", pady=15)

        self.lbl_year = ctk.CTkLabel(self.param_frame, text="Active Tax Year:", font=ctk.CTkFont(weight="bold"))
        self.lbl_year.pack(side="left", padx=(0, 10))
        self.entry_year = ctk.CTkEntry(self.param_frame, width=100, justify="center")
        self.entry_year.pack(side="left", padx=(0, 30))

        self.lbl_quarter = ctk.CTkLabel(self.param_frame, text="Active Quarter Target:",
                                        font=ctk.CTkFont(weight="bold"))
        self.lbl_quarter.pack(side="left", padx=(0, 10))
        self.entry_quarter = ctk.CTkComboBox(self.param_frame, values=["1", "2", "3", "4"], width=80, justify="center")
        self.entry_quarter.pack(side="left")

        self.btn_run = ctk.CTkButton(
            self, text="⚡ Run Supplemental Calculations", height=45,
            font=ctk.CTkFont(size=15, weight="bold"), command=self.start_processing_thread
        )
        self.btn_run.grid(row=2, column=0, pady=(15, 10), padx=40, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self, width=620, mode="indeterminate")
        self.progress_bar.grid(row=3, column=0, pady=5)
        self.progress_bar.set(0)

        self.real_file_path.trace_add("write", lambda *args: self.save_app_config())
        self.pp_file_path.trace_add("write", lambda *args: self.save_app_config())
        self.output_dir_path.trace_add("write", lambda *args: self.save_app_config())

        self.load_app_config()

    def change_appearance_theme(self, selection_mode):
        ctk.set_appearance_mode(selection_mode)
        self.save_app_config()

    def load_app_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    cfg = json.load(f)
                    self.real_file_path.set(cfg.get("real_file_path", ""))
                    self.pp_file_path.set(cfg.get("pp_file_path", ""))
                    self.output_dir_path.set(cfg.get("output_dir_path", ""))
                    self.entry_year.insert(0, str(cfg.get("tax_year", "2026")))
                    self.entry_quarter.set(str(cfg.get("quarter", "2")))
                    self.theme_menu.set(cfg.get("appearance_mode", "System"))
                    return
            except Exception:
                pass
        self.entry_year.insert(0, "2026")
        self.entry_quarter.set("2")
        self.theme_menu.set("System")

    def save_app_config(self):
        try:
            cfg = {
                "real_file_path": self.real_file_path.get(),
                "pp_file_path": self.pp_file_path.get(),
                "output_dir_path": self.output_dir_path.get(),
                "tax_year": self.entry_year.get(),
                "quarter": self.entry_quarter.get(),
                "appearance_mode": self.theme_menu.get()
            }
            with open(CONFIG_FILE, "w") as f:
                json.dump(cfg, f, indent=4)
        except Exception:
            pass

    def browse_real_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Documents", "*.xlsx *.xls")])
        if path:
            self.real_file_path.set(path)

    def browse_pp_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Documents", "*.xlsx *.xls")])
        if path:
            self.pp_file_path.set(path)

    def browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir_path.set(path)

    def start_processing_thread(self):
        real = self.real_file_path.get().strip()
        pp = self.pp_file_path.get().strip()
        out = self.output_dir_path.get().strip()

        if not real and not pp:
            AnimatedDialog(self, "Missing Source Files", "Please provide at least one source file to execute.", "error")
            return
        if not out:
            AnimatedDialog(self, "Missing Output Target", "An output directory path.py is completely required.", "error")
            return
        try:
            tax_year = int(self.entry_year.get().strip())
            quarter = int(self.entry_quarter.get().strip())
        except ValueError:
            AnimatedDialog(self, "Invalid Inputs", "Tax Year and Quarter must be valid numerical parameters.", "error")
            return

        self.save_app_config()
        self.btn_run.configure(state="disabled", text="🔄 Processing Sheets... Please Wait")
        self.progress_bar.start()

        worker = threading.Thread(target=self.execute_processing_pipeline, args=(real, pp, out, tax_year, quarter),
                                  daemon=True)
        worker.start()

    def execute_processing_pipeline(self, real_path, pp_path, output_dir, tax_year, quarter):
        try:
            # ------------------------------------------------------------------
            # 🛠️ YOUR ENGINE RUNNER INTEGRATION POINT (Call excel_engine calculations here)
            # ------------------------------------------------------------------
            time.sleep(3.5)
            self.after(0, self.handle_pipeline_success, output_dir)
        except Exception as err:
            self.after(0, self.handle_pipeline_failure, str(err))

    def handle_pipeline_success(self, output_dir):
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self.btn_run.configure(state="normal", text="⚡ Run Supplemental Calculations")
        AnimatedDialog(self, "Generation Successful",
                       f"Consolidated supplementary tax roll workbooks completely initialized at:\n{output_dir}",
                       mode="success")

    def handle_pipeline_failure(self, error_message):
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self.btn_run.configure(state="normal", text="⚡ Run Supplemental Calculations")
        AnimatedDialog(self, "Execution Runtime Crash",
                       f"An absolute exception halted processing loops:\n{error_message}", mode="error")
