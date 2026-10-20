import os
import sys
import json
import queue
import logging
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

from core import APP_NAME, __version__
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

LOG_FILE = os.path.join(os.path.expanduser("~"), "mouseframe_organizer.log")
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), "mouseframe_organizer_settings.json")

ctk.set_default_color_theme("blue")

SECTION_FONT = dict(size=15, weight="bold")
HINT_COLOR = ("gray40", "gray60")


def _clean_path(text):
    """Strips whitespace and quotes that Windows 'Copy as path' adds."""
    return text.strip().strip('"').strip("'")


def _load_settings():
    """Last used paths and theme; {} when the file is missing or unreadable."""
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


class MouseframeOrganizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self._settings = _load_settings()
        ctk.set_appearance_mode(self._settings.get("theme", "System"))

        self.title(f"{APP_NAME} {__version__}")
        self.geometry("980x820")
        self.minsize(900, 740)

        # Worker threads must never touch Tk directly; they queue callables run on the UI thread
        self._ui_calls = queue.SimpleQueue()
        self._last_output = None

        self._create_header()
        self._create_footer()  # packed before the tabs so it stays visible at the bottom
        self._create_tabs()
        self._restore_paths()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
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

    def _create_header(self):
        header = ctk.CTkFrame(self, corner_radius=10)
        header.pack(fill="x", padx=15, pady=(15, 5))

        title_container = ctk.CTkFrame(header, fg_color="transparent")
        title_container.pack(side="left", fill="both", expand=True, padx=15, pady=10)
        ctk.CTkLabel(title_container, text=APP_NAME, font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            title_container,
            text=f"Version {__version__}  ·  Organize raw MouseFrame gait workbooks into one Excel file "
                 "with per-animal means and group statistics.",
            font=ctk.CTkFont(size=12),
            text_color=HINT_COLOR
        ).pack(anchor="w")

        theme_container = ctk.CTkFrame(header, fg_color="transparent")
        theme_container.pack(side="right", padx=15, pady=10)
        ctk.CTkLabel(theme_container, text="Theme").pack(side="left", padx=(0, 8))
        self.theme_dropdown = ctk.CTkOptionMenu(
            theme_container,
            values=["System", "Light", "Dark"],
            width=95,
            command=self._on_theme_change
        )
        self.theme_dropdown.set(self._settings.get("theme", "System"))
        self.theme_dropdown.pack(side="left")

    def _on_theme_change(self, new_mode):
        ctk.set_appearance_mode(new_mode)
        # get_appearance_mode() resolves "System" to the actual "Light"/"Dark"
        self.preview_frame.apply_theme_colors(ctk.get_appearance_mode())

    def _create_footer(self):
        footer = ctk.CTkFrame(self, corner_radius=10)
        footer.pack(side="bottom", fill="x", padx=15, pady=(5, 15))

        self.progress_bar = ctk.CTkProgressBar(footer)
        self.progress_bar.pack(fill="x", padx=15, pady=(12, 6))
        self.progress_bar.set(0)

        row = ctk.CTkFrame(footer, fg_color="transparent")
        row.pack(fill="x", padx=15, pady=(0, 10))
        self.btn_open_folder = ctk.CTkButton(row, text="Open folder", width=110, state="disabled",
                                             command=lambda: self._open_path(os.path.dirname(self._last_output)))
        self.btn_open_folder.pack(side="right")
        self.btn_open_file = ctk.CTkButton(row, text="Open file", width=110, state="disabled",
                                           command=lambda: self._open_path(self._last_output))
        self.btn_open_file.pack(side="right", padx=(0, 8))
        self.status_label = ctk.CTkLabel(row, text="Ready.", anchor="w", justify="left")
        self.status_label.pack(side="left", fill="x", expand=True)

    def _create_tabs(self):
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=15, pady=5)

        self.tab_organizer = self.tabview.add("Organizer")
        self.tab_tutorial = self.tabview.add("User guide")

        self._build_organizer_tab()
        TutorialFrame(self.tab_tutorial).pack(fill="both", expand=True, padx=5, pady=5)

    def _build_organizer_tab(self):
        tab = self.tab_organizer

        # 1. Files
        files = ctk.CTkFrame(tab, fg_color="transparent")
        files.pack(fill="x", padx=10, pady=(0, 4))
        files.columnconfigure(1, weight=1)
        ctk.CTkLabel(files, text="1. Files", font=ctk.CTkFont(**SECTION_FONT)).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))

        self.input_entry = self._file_row(files, 1, "MouseFrame workbook", "Raw export (.xlsx or .xlsm)",
                                          lambda: self._browse(self.input_entry, filedialog.askopenfilename,
                                                               title="Select raw MouseFrame workbook",
                                                               filetypes=[("Excel workbooks", "*.xlsx *.xlsm")]))
        self.mapping_entry = self._file_row(files, 2, "Mapping file (optional)",
                                            "CSV or Excel with columns ID, Genotype, Sex",
                                            lambda: self._browse(self.mapping_entry, filedialog.askopenfilename,
                                                                 title="Select mapping file",
                                                                 filetypes=[("CSV or Excel", "*.csv *.xlsx *.xlsm"),
                                                                            ("All files", "*.*")]))
        ctk.CTkButton(files, text="✕", width=32, fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"), command=lambda: self.mapping_entry.delete(0, "end")).grid(
            row=2, column=3, padx=(6, 0), pady=4)
        self.output_entry = self._file_row(files, 3, "Output folder (optional)", "Default: folder of the workbook",
                                           lambda: self._browse(self.output_entry, filedialog.askdirectory,
                                                                title="Select output folder"))

        # 2. Options
        options = ctk.CTkFrame(tab, fg_color="transparent")
        options.pack(fill="x", padx=10, pady=4)
        options.columnconfigure(2, weight=1)
        ctk.CTkLabel(options, text="2. Options", font=ctk.CTkFont(**SECTION_FONT)).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ctk.CTkLabel(options, text="Group colors").grid(row=1, column=0, sticky="w", pady=4)
        self.visual_mode_dropdown = ctk.CTkOptionMenu(options, values=VISUAL_MODES, width=200)
        self.visual_mode_dropdown.set(MODE_PASTEL_ACCENTS)
        self.visual_mode_dropdown.grid(row=1, column=1, sticky="w", padx=10, pady=4)
        ctk.CTkLabel(options, text="Colors need a mapping file", text_color=HINT_COLOR).grid(
            row=1, column=2, sticky="w", pady=4)

        ctk.CTkLabel(options, text="Decimal-fix threshold").grid(row=2, column=0, sticky="w", pady=4)
        self.threshold_entry = ctk.CTkEntry(options, width=200)
        self.threshold_entry.insert(0, "500")
        self.threshold_entry.grid(row=2, column=1, sticky="w", padx=10, pady=4)
        ctk.CTkLabel(options, text="Values above this are divided by 1000 (2.456 read as 2456)",
                     text_color=HINT_COLOR).grid(row=2, column=2, sticky="w", pady=4)

        self.chk_summary_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(options, text="Add summary sheets", variable=self.chk_summary_var).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 4))
        ctk.CTkLabel(options, text="Mean per animal; N / Mean / SEM per genotype × sex with a mapping file",
                     text_color=HINT_COLOR).grid(row=3, column=2, sticky="w", pady=(8, 4))

        # 3. Actions
        actions = ctk.CTkFrame(tab, fg_color="transparent")
        actions.pack(fill="x", padx=10, pady=(8, 6))
        self.btn_process = ctk.CTkButton(actions, text="Export workbook", width=170, height=36,
                                         font=ctk.CTkFont(size=14, weight="bold"), command=self._on_process_click)
        self.btn_process.pack(side="right")
        self.btn_preview = ctk.CTkButton(actions, text="Preview", width=120, height=36, fg_color="transparent",
                                         border_width=2, text_color=("gray10", "gray90"),
                                         command=self._on_preview_click)
        self.btn_preview.pack(side="right", padx=(0, 10))

        # 4. Preview (tables and warnings)
        self.preview_frame = PreviewTableFrame(tab)
        self.preview_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))

    def _file_row(self, parent, row, label, placeholder, browse):
        """Label, entry and Browse button on one grid row. Returns the entry."""
        ctk.CTkLabel(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        entry = ctk.CTkEntry(parent, placeholder_text=placeholder)
        entry.grid(row=row, column=1, sticky="ew", padx=10, pady=4)
        ctk.CTkButton(parent, text="Browse…", width=100, command=browse).grid(row=row, column=2, pady=4)
        return entry

    def _browse(self, entry, ask, **options):
        """Opens a file/folder dialog starting next to the current path and writes the choice into entry."""
        current = _clean_path(entry.get()) or _clean_path(self.input_entry.get())
        if current:
            options["initialdir"] = current if os.path.isdir(current) else os.path.dirname(current)
        path = ask(**options)
        if path:
            entry.delete(0, "end")
            entry.insert(0, os.path.normpath(path))

    def _restore_paths(self):
        for key, entry in self._path_entries():
            path = self._settings.get(key, "")
            if path and os.path.exists(path):
                entry.insert(0, path)

    def _path_entries(self):
        return (("input_file", self.input_entry), ("mapping_file", self.mapping_entry),
                ("output_dir", self.output_entry))

    def _on_close(self):
        """Remembers paths and theme for the next start, then closes."""
        settings = {key: _clean_path(entry.get()) for key, entry in self._path_entries()}
        settings["theme"] = self.theme_dropdown.get()
        try:
            with open(SETTINGS_FILE + ".part", "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2)
            os.replace(SETTINGS_FILE + ".part", SETTINGS_FILE)
        except OSError:
            log.warning("Could not save settings to %s", SETTINGS_FILE, exc_info=True)
        self.destroy()

    def _log_status(self, message, progress_pct=None):
        """Thread-safe status update (worker threads must not touch widgets directly)."""
        def _update():
            self.status_label.configure(text=message)
            if progress_pct is not None:
                self.progress_bar.set(progress_pct / 100.0)

        self._call_ui(_update)

    def _set_busy(self, busy, indeterminate=False):
        state = "disabled" if busy else "normal"
        self.btn_process.configure(state=state)
        self.btn_preview.configure(state=state)
        if busy and indeterminate:
            self.progress_bar.configure(mode="indeterminate")
            self.progress_bar.start()
        elif not busy:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")

    def _read_inputs(self):
        """Validates the form. Returns a dict of inputs, or None after showing an error."""
        input_file = _clean_path(self.input_entry.get())
        mapping_file = _clean_path(self.mapping_entry.get())
        output_dir = _clean_path(self.output_entry.get())

        if not input_file or not os.path.isfile(input_file):
            messagebox.showerror("Workbook missing", "Select the raw MouseFrame workbook (.xlsx) first.")
            return None
        if mapping_file and not os.path.isfile(mapping_file):
            messagebox.showerror("Mapping file not found", f"This mapping file does not exist:\n{mapping_file}")
            return None
        if output_dir and not os.path.isdir(output_dir):
            messagebox.showerror("Output folder not found", f"This output folder does not exist:\n{output_dir}")
            return None
        try:
            threshold = float(self.threshold_entry.get().strip().replace(",", "."))
            if threshold <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid threshold", "The decimal-fix threshold must be a positive number, "
                                                      "for example 500.")
            return None

        return {
            "input_file": input_file,
            "mapping_file": mapping_file or None,
            "output_dir": output_dir or None,
            "threshold": threshold,
        }

    def _run_in_background(self, work, on_success, error_title, indeterminate=False):
        """Runs work() in a thread; calls on_success(result) or shows an error, on the UI thread."""
        self._set_busy(True, indeterminate)

        def _worker():
            try:
                result = work()
            except Exception as e:
                log.exception(error_title)
                msg = str(e) or e.__class__.__name__  # capture now: 'e' is cleared after the except block

                def _fail_ui():
                    self._set_busy(False)
                    self._log_status(f"{error_title}.", 0)
                    messagebox.showerror(error_title, f"{msg}\n\nTechnical details are saved in:\n{LOG_FILE}")

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

        self._log_status("Reading workbook…")

        def _work():
            data = find_table_data(inputs["input_file"], threshold=inputs["threshold"])
            mapping_dict, _ = load_mapping_file(inputs["mapping_file"])
            return build_records(data, mapping_dict, inputs["threshold"]), bool(mapping_dict)

        def _show(result):
            (records, warnings), has_mapping = result
            self.preview_frame.update_data(records, warnings, has_mapping)
            sheets = len({r["Source Sheet"] for r in records})
            text = f"Preview: {len(records)} tables in {sheets} sheets."
            if warnings:
                text += f" {len(warnings)} warning(s), see the Warnings view."
            self._log_status(text, 100)

        self._run_in_background(_work, _show, "Could not read the workbook", indeterminate=True)

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
            self._last_output = out_path
            self.btn_open_file.configure(state="normal")
            self.btn_open_folder.configure(state="normal")
            note = f"\n\n{len(warnings)} warning(s) are listed in the Processing_Log sheet." if warnings else ""
            if messagebox.askyesno("Export finished", f"Workbook saved:\n{out_path}{note}\n\nOpen it now?"):
                self._open_path(out_path)

        self._run_in_background(_work, _done, "Export failed")

    @staticmethod
    def _open_path(path):
        """Opens a file or folder with the system default application."""
        try:
            if sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        except OSError as e:
            messagebox.showwarning("Could not open", f"Could not open:\n{path}\n\n{e}")
