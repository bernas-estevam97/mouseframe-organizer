import os
import sys
import queue
import logging
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

from core.parser import find_table_data
from core.mapping import load_mapping_file
from core.excel_writer import (
    build_records,
    process_and_export_file,
    VISUAL_MODES,
    MODE_PASTEL_ACCENTS
)
from gui.preview_table import PreviewTableFrame
from gui.tutorial import TutorialFrame

log = logging.getLogger(__name__)

ctk.set_appearance_mode("Dark")  # Default mode: "Dark", can be switched to "Light" or "System"
ctk.set_default_color_theme("blue")


def _clean_path(text):
    """Strips whitespace and quotes that Windows 'Copy as path' adds."""
    return text.strip().strip('"').strip("'")


class MouseframeOrganizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Mouseframe Data Organizer")
        self.geometry("980x820")
        self.minsize(850, 680)

        # Worker threads must never touch Tk directly; they queue callables run on the UI thread
        self._ui_calls = queue.SimpleQueue()

        self._create_header_and_theme_bar()
        self._create_tabs()
        self._drain_ui_calls()

    def _call_ui(self, fn):
        """Thread-safe: schedules fn() on the UI thread."""
        self._ui_calls.put(fn)

    def _drain_ui_calls(self):
        try:
            while not self._ui_calls.empty():  # single consumer, so empty() is reliable here
                self._ui_calls.get()()
        finally:
            self.after(50, self._drain_ui_calls)  # keep polling even if a callback raised

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
            # get_appearance_mode() resolves "System" to the actual "Light"/"Dark"
            self.preview_frame.apply_theme_colors(ctk.get_appearance_mode())

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
            text="Generate Summary Sheets (mean per animal; N / Mean / SEM per Genotype × Sex when a mapping file is set)",
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

        self.status_label = ctk.CTkLabel(self.log_group, text="Ready.", anchor="w", justify="left", wraplength=850)
        self.status_label.pack(fill="x", padx=15, pady=(0, 10))

    def _browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Raw Mouseframe Excel File",
            filetypes=[("Excel Workbooks", "*.xlsx *.xlsm")]
        )
        if file_path:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, file_path)

    def _browse_mapping_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Metadata Mapping File",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
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
        """Thread-safe status update (worker threads must not touch widgets directly)."""
        def _update():
            self.status_label.configure(text=message)
            if progress_pct is not None:
                self.progress_bar.set(progress_pct / 100.0)

        self._call_ui(_update)

    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        self.btn_process.configure(state=state)
        self.btn_preview.configure(state=state)

    def _read_inputs(self):
        """Validates the form. Returns a dict of inputs, or None after showing an error."""
        input_file = _clean_path(self.input_entry.get())
        mapping_file = _clean_path(self.mapping_entry.get())
        output_dir = _clean_path(self.output_entry.get())

        if not input_file or not os.path.isfile(input_file):
            messagebox.showerror("Error", "Please select a valid raw Mouseframe Excel file.")
            return None
        if mapping_file and not os.path.isfile(mapping_file):
            messagebox.showerror("Error", f"Mapping file not found:\n{mapping_file}")
            return None
        if output_dir and not os.path.isdir(output_dir):
            messagebox.showerror("Error", f"Output directory does not exist:\n{output_dir}")
            return None
        try:
            threshold = float(self.threshold_entry.get().strip().replace(",", "."))
            if threshold <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter a positive numeric threshold.")
            return None

        return {
            "input_file": input_file,
            "mapping_file": mapping_file or None,
            "output_dir": output_dir or None,
            "threshold": threshold,
        }

    def _run_in_background(self, work, on_success, error_title):
        """Runs work() in a thread; calls on_success(result) or shows an error, on the UI thread."""
        self._set_busy(True)

        def _worker():
            try:
                result = work()
            except Exception as e:
                log.exception(error_title)
                msg = str(e) or e.__class__.__name__  # capture now: 'e' is cleared after the except block

                def _fail_ui():
                    self._set_busy(False)
                    self._log_status(f"{error_title}.", 0)
                    messagebox.showerror(error_title, msg)

                self._call_ui(_fail_ui)
                return

            def _success_ui():
                self._set_busy(False)
                on_success(result)

            self._call_ui(_success_ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_preview_click(self):
        inputs = self._read_inputs()
        if not inputs:
            return

        self._log_status("Loading and parsing raw data preview...", 20)

        def _work():
            data = find_table_data(inputs["input_file"], threshold=inputs["threshold"])
            mapping_dict, _ = load_mapping_file(inputs["mapping_file"])
            return build_records(data, mapping_dict, inputs["threshold"]), bool(mapping_dict)

        def _show(result):
            (records, warnings), has_mapping = result
            self.preview_frame.update_data(records, has_mapping)
            sheets = len({r["Source Sheet"] for r in records})
            text = f"Preview updated. {len(records)} tables parsed across {sheets} sheets."
            if warnings:
                text += f" {len(warnings)} warning(s):\n" + "\n".join(f"• {w}" for w in warnings[:5])
                if len(warnings) > 5:
                    text += f"\n… see the Processing_Log sheet after export for all {len(warnings)}."
            self._log_status(text, 100)

        self._run_in_background(_work, _show, "Could not parse file")

    def _on_process_click(self):
        inputs = self._read_inputs()
        if not inputs:
            return

        visual_mode = self.visual_mode_dropdown.get()
        generate_summary = self.chk_summary_var.get()

        def _work():
            return process_and_export_file(
                visual_mode=visual_mode,
                generate_summary=generate_summary,
                progress_callback=self._log_status,
                **inputs
            )

        def _done(result):
            out_path, warnings = result
            note = f"\n\n{len(warnings)} warning(s) listed in the Processing_Log sheet." if warnings else ""
            if messagebox.askyesno(
                "Success!",
                f"Organized Excel workbook saved successfully:\n\n{out_path}{note}\n\nDo you want to open the output folder now?"
            ):
                self._open_folder(os.path.dirname(out_path))

        self._run_in_background(_work, _done, "Processing failed")

    @staticmethod
    def _open_folder(folder):
        try:
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", folder])
        except OSError as e:
            messagebox.showwarning("Open Folder", f"Could not open the folder:\n{e}")
