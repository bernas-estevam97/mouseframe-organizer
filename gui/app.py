import os
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

from core.parser import find_table_data
from core.mapping import load_mapping_file
from core.excel_writer import (
    process_and_export_file,
    VISUAL_MODES,
    MODE_PASTEL_ACCENTS
)
from gui.preview_table import PreviewTableFrame
from gui.tutorial import TutorialFrame

ctk.set_appearance_mode("Dark")  # Default mode: "Dark", can be switched to "Light" or "System"
ctk.set_default_color_theme("blue")


class MouseframeOrganizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Mouseframe Data Organizer")
        self.geometry("980, 820")
        self.minsize(850, 680)

        self.input_file_path = ""
        self.mapping_file_path = ""
        self.output_dir_path = ""

        self._parsed_data_cache = None
        self._mapping_dict_cache = {}

        self._create_header_and_theme_bar()
        self._create_tabs()

    def _create_header_and_theme_bar(self):
        # Header Section with Title and Theme Switcher
        self.header_frame = ctk.CTkFrame(self, corner_radius=10)
        self.header_frame.pack(fill="x", padx=15, pady=(15, 5))

        title_container = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        title_container.pack(side="left", fill="both", expand=True, padx=15, pady=10)

        self.title_label = ctk.CTkLabel(
            title_container,
            text="Mouseframe Gait Data Organizer",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.title_label.pack(anchor="w")

        self.subtitle_label = ctk.CTkLabel(
            title_container,
            text="Clean, group, and format raw Mouseframe gait tables into structured Excel sheets with automated statistics.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.subtitle_label.pack(anchor="w")

        # Theme Switcher Control
        theme_container = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        theme_container.pack(side="right", padx=15, pady=10)

        ctk.CTkLabel(
            theme_container,
            text="Theme Mode:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=(0, 8))

        self.theme_dropdown = ctk.CTkOptionMenu(
            theme_container,
            values=["Dark", "Light", "System"],
            width=95,
            command=self._on_theme_change
        )
        self.theme_dropdown.set("Dark")
        self.theme_dropdown.pack(side="left")

    def _on_theme_change(self, new_mode):
        ctk.set_appearance_mode(new_mode)
        if hasattr(self, 'preview_frame'):
            self.preview_frame.apply_theme_colors(new_mode)

    def _create_tabs(self):
        # Tabview navigation container
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        self.tab_organizer = self.tabview.add("🚀 Organizer & Export")
        self.tab_tutorial = self.tabview.add("📖 Tutorial & Guide")

        self._build_organizer_tab()
        self._build_tutorial_tab()

    def _build_tutorial_tab(self):
        self.tutorial_frame = TutorialFrame(self.tab_tutorial)
        self.tutorial_frame.pack(fill="both", expand=True, padx=5, pady=5)

    def _build_organizer_tab(self):
        # Scrollable main container inside tab
        self.main_container = ctk.CTkScrollableFrame(self.tab_organizer, corner_radius=10)
        self.main_container.pack(fill="both", expand=True, padx=5, pady=5)

        # 1. File Selection Section
        self.files_group = ctk.CTkFrame(self.main_container)
        self.files_group.pack(fill="x", padx=10, pady=10)

        self.files_label = ctk.CTkLabel(
            self.files_group,
            text="1. File Inputs & Outputs",
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.files_label.grid(row=0, column=0, columnspan=3, sticky="w", padx=15, pady=(10, 10))

        # Raw Excel File Input
        ctk.CTkLabel(self.files_group, text="Raw Mouseframe Excel File:").grid(row=1, column=0, sticky="w", padx=15, pady=5)
        self.input_entry = ctk.CTkEntry(self.files_group, placeholder_text="Select raw data Excel file (.xlsx)...")
        self.input_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        self.btn_browse_input = ctk.CTkButton(self.files_group, text="Browse...", width=100, command=self._browse_input_file)
        self.btn_browse_input.grid(row=1, column=2, padx=(0, 15), pady=5)

        # Optional Mapping File Input
        ctk.CTkLabel(self.files_group, text="Metadata Mapping File (Optional):").grid(row=2, column=0, sticky="w", padx=15, pady=5)
        self.mapping_entry = ctk.CTkEntry(self.files_group, placeholder_text="Select CSV or Excel mapping file (ID, Genotype, Sex)...")
        self.mapping_entry.grid(row=2, column=1, sticky="ew", padx=10, pady=5)
        
        mapping_btns_frame = ctk.CTkFrame(self.files_group, fg_color="transparent")
        mapping_btns_frame.grid(row=2, column=2, padx=(0, 15), pady=5)
        self.btn_browse_mapping = ctk.CTkButton(mapping_btns_frame, text="Browse...", width=70, command=self._browse_mapping_file)
        self.btn_browse_mapping.pack(side="left", padx=(0, 5))
        self.btn_clear_mapping = ctk.CTkButton(mapping_btns_frame, text="Clear", width=50, fg_color="gray", command=self._clear_mapping_file)
        self.btn_clear_mapping.pack(side="left")

        # Output Folder Option
        ctk.CTkLabel(self.files_group, text="Output Directory (Optional):").grid(row=3, column=0, sticky="w", padx=15, pady=5)
        self.output_entry = ctk.CTkEntry(self.files_group, placeholder_text="Default: Same folder as input file")
        self.output_entry.grid(row=3, column=1, sticky="ew", padx=10, pady=(5, 15))
        self.btn_browse_output = ctk.CTkButton(self.files_group, text="Browse...", width=100, command=self._browse_output_dir)
        self.btn_browse_output.grid(row=3, column=2, padx=(0, 15), pady=(5, 15))

        self.files_group.columnconfigure(1, weight=1)

        # 2. Options & Group Representation Section
        self.options_group = ctk.CTkFrame(self.main_container)
        self.options_group.pack(fill="x", padx=10, pady=10)

        self.options_label = ctk.CTkLabel(
            self.options_group,
            text="2. Group Representation & Processing Options",
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.options_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(10, 10))

        # Visual Grouping Dropdown
        ctk.CTkLabel(self.options_group, text="Group Visual Style:").grid(row=1, column=0, sticky="w", padx=15, pady=5)
        self.visual_mode_dropdown = ctk.CTkOptionMenu(
            self.options_group,
            values=VISUAL_MODES,
            width=320
        )
        self.visual_mode_dropdown.set(MODE_PASTEL_ACCENTS)
        self.visual_mode_dropdown.grid(row=1, column=1, sticky="w", padx=10, pady=5)

        # Threshold input
        ctk.CTkLabel(self.options_group, text="Unit Conversion Threshold:").grid(row=2, column=0, sticky="w", padx=15, pady=5)
        self.threshold_entry = ctk.CTkEntry(self.options_group, width=100)
        self.threshold_entry.insert(0, "500.0")
        self.threshold_entry.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        ctk.CTkLabel(self.options_group, text="(Values > threshold are divided by 1000.0)", text_color="gray").grid(row=2, column=1, sticky="w", padx=(120, 10), pady=5)

        # Summary sheet checkbox
        self.chk_summary_var = ctk.BooleanVar(value=True)
        self.chk_summary = ctk.CTkCheckBox(
            self.options_group,
            text="Generate Automated Summary Statistics Sheet (Mean ± SEM per Genotype/Sex)",
            variable=self.chk_summary_var,
            font=ctk.CTkFont(weight="bold")
        )
        self.chk_summary.grid(row=3, column=0, columnspan=2, sticky="w", padx=15, pady=(10, 15))

        # 3. Actions Frame
        self.actions_group = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.actions_group.pack(fill="x", padx=10, pady=5)

        self.btn_preview = ctk.CTkButton(
            self.actions_group,
            text="🔍 Preview Data",
            width=140,
            fg_color="#27AE60",
            hover_color="#219150",
            command=self._on_preview_click
        )
        self.btn_preview.pack(side="left", padx=(5, 10), pady=5)

        self.btn_process = ctk.CTkButton(
            self.actions_group,
            text="🚀 Process & Export Excel",
            width=200,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2980B9",
            hover_color="#1F618D",
            command=self._on_process_click
        )
        self.btn_process.pack(side="left", padx=5, pady=5)

        # 4. Preview Table Frame
        self.preview_frame = PreviewTableFrame(self.main_container)
        self.preview_frame.pack(fill="x", padx=10, pady=10)

        # 5. Log & Progress Console
        self.log_group = ctk.CTkFrame(self.main_container)
        self.log_group.pack(fill="x", padx=10, pady=10)

        self.progress_bar = ctk.CTkProgressBar(self.log_group)
        self.progress_bar.pack(fill="x", padx=15, pady=(15, 5))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(self.log_group, text="Ready.", anchor="w")
        self.status_label.pack(fill="x", padx=15, pady=(0, 10))

    def _browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Raw Mouseframe Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, file_path)
            self._parsed_data_cache = None

    def _browse_mapping_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Metadata Mapping File",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if file_path:
            self.mapping_entry.delete(0, "end")
            self.mapping_entry.insert(0, file_path)

    def _clear_mapping_file(self):
        self.mapping_entry.delete(0, "end")

    def _browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, dir_path)

    def _log_status(self, message, progress_pct=None):
        def _update():
            self.status_label.configure(text=message)
            if progress_pct is not None:
                self.progress_bar.set(progress_pct / 100.0)

        self.after(0, _update)

    def _on_preview_click(self):
        input_file = self.input_entry.get().strip().strip('"').strip("'")
        mapping_file = self.mapping_entry.get().strip().strip('"').strip("'")

        if not input_file or not os.path.isfile(input_file):
            messagebox.showerror("Error", "Please select a valid raw Mouseframe Excel file.")
            return

        try:
            threshold = float(self.threshold_entry.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid numeric threshold.")
            return

        self._log_status("Loading and parsing raw data preview...", 20)

        def _preview_worker():
            try:
                data = find_table_data(input_file, threshold=threshold)
                mapping_dict = {}
                if mapping_file and os.path.isfile(mapping_file):
                    mapping_dict, _, _ = load_mapping_file(mapping_file)

                def _ui_update():
                    self.preview_frame.update_data(data, mapping_dict)
                    self._log_status(f"Preview updated. {len(data)} tables parsed across {len(set(d.get('Sheet_Source') for d in data))} sheets.", 100)

                self.after(0, _ui_update)
            except Exception as e:
                def _error_update():
                    messagebox.showerror("Parsing Error", f"Could not parse file:\n{str(e)}")
                    self._log_status("Error loading preview.", 0)

                self.after(0, _error_update)

        threading.Thread(target=_preview_worker, daemon=True).start()

    def _on_process_click(self):
        input_file = self.input_entry.get().strip().strip('"').strip("'")
        mapping_file = self.mapping_entry.get().strip().strip('"').strip("'")
        output_dir = self.output_entry.get().strip().strip('"').strip("'")

        if not input_file or not os.path.isfile(input_file):
            messagebox.showerror("Error", "Please select a valid raw Mouseframe Excel file.")
            return

        try:
            threshold = float(self.threshold_entry.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid numeric threshold.")
            return

        visual_mode = self.visual_mode_dropdown.get()
        generate_summary = self.chk_summary_var.get()

        self.btn_process.configure(state="disabled")
        self.btn_preview.configure(state="disabled")

        def _process_worker():
            try:
                out_path = process_and_export_file(
                    input_file=input_file,
                    mapping_file=mapping_file if mapping_file else None,
                    output_dir=output_dir if output_dir else None,
                    threshold=threshold,
                    visual_mode=visual_mode,
                    generate_summary=generate_summary,
                    progress_callback=self._log_status
                )

                def _success_ui():
                    self.btn_process.configure(state="normal")
                    self.btn_preview.configure(state="normal")
                    
                    res = messagebox.askyesno(
                        "Success!",
                        f"Organized Excel workbook saved successfully:\n\n{out_path}\n\nDo you want to open the output folder now?"
                    )
                    if res:
                        folder = os.path.dirname(out_path)
                        if os.name == 'nt':
                            os.startfile(folder)
                        else:
                            subprocess.Popen(['open' if os.name == 'posix' else 'xdg-open', folder])

                self.after(0, _success_ui)

            except Exception as e:
                def _fail_ui():
                    self.btn_process.configure(state="normal")
                    self.btn_preview.configure(state="normal")
                    messagebox.showerror("Processing Failed", f"An error occurred during processing:\n\n{str(e)}")
                    self._log_status("Processing failed.", 0)

                self.after(0, _fail_ui)

        threading.Thread(target=_process_worker, daemon=True).start()
