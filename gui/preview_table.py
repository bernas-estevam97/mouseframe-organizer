import customtkinter as ctk
from tkinter import ttk

from core.parser import sort_tags, METRIC_COLUMNS
from core.excel_writer import UNMAPPED

TABLES_VIEW = "Tables"


class PreviewTableFrame(ctk.CTkFrame):
    """
    Preview of parsed tables (sortable ttk Treeview) and a Warnings view, switched by a segmented button.
    Rows are green when mapped, red when unmapped and amber when metrics are missing.
    """

    # (key, heading, width, anchor)
    COLUMNS = [
        ("table_id", "Table ID", 170, "w"),
        ("animal_id", "Animal ID", 100, "w"),
        ("timepoint", "Timepoint", 80, "center"),
        ("week", "Week", 70, "center"),
        ("genotype", "Genotype", 100, "center"),
        ("sex", "Sex", 70, "center"),
        ("runs", "Runs", 50, "center"),
        ("metrics", "Metrics", 60, "center"),
    ]

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._rows = []
        self._sort_key, self._sort_reverse = "timepoint", False

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(bar, text="Preview", font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
        self.view_switch = ctk.CTkSegmentedButton(bar, values=[TABLES_VIEW, "Warnings"], command=self._show_view)
        self.view_switch.set(TABLES_VIEW)
        self.view_switch.pack(side="left", padx=15)
        self.only_unmapped_var = ctk.BooleanVar(value=False)
        self.only_unmapped = ctk.CTkCheckBox(bar, text="Only unmapped", variable=self.only_unmapped_var,
                                             command=self._refresh, state="disabled")
        self.only_unmapped.pack(side="right")

        self.table_container = ctk.CTkFrame(self, fg_color="transparent")
        self.table_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # "clam" applies app-wide to ttk widgets; it is the only built-in theme that honors custom colors
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # no outer border

        self.tree = ttk.Treeview(self.table_container, columns=[c[0] for c in self.COLUMNS],
                                 show="headings", height=6)
        for key, heading, width, anchor in self.COLUMNS:
            self.tree.heading(key, text=heading, command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=width, anchor=anchor)

        v_scrollbar = ctk.CTkScrollbar(self.table_container, orientation="vertical", command=self.tree.yview)
        h_scrollbar = ctk.CTkScrollbar(self.table_container, orientation="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.warnings_box = ctk.CTkTextbox(self, wrap="word")
        self._set_warnings([])

        self.apply_theme_colors(ctk.get_appearance_mode())

    def apply_theme_colors(self, mode="Dark"):
        """Configures Treeview colors for the Light/Dark appearance mode and the current UI scaling."""
        is_dark = mode.lower() == "dark"
        # ttk does not follow customtkinter's DPI scaling, so scale the pixel row height by hand
        scaling = ctk.ScalingTracker.get_widget_scaling(self)

        bg = "#2B2B2B" if is_dark else "#FFFFFF"
        fg = "#E8E8E8" if is_dark else "#1A1A1A"
        self.style.configure("Treeview", background=bg, foreground=fg, fieldbackground=bg, borderwidth=0,
                             rowheight=int(24 * scaling), font=("Segoe UI", 10))
        self.style.map("Treeview", background=[("selected", "#1F6AA5")], foreground=[("selected", "#FFFFFF")])
        self.style.configure("Treeview.Heading", background="#333333" if is_dark else "#E6E6E6",
                             foreground=fg, relief="flat", borderwidth=0, font=("Segoe UI", 10, "bold"))
        self.style.map("Treeview.Heading", background=[("active", "#3D3D3D" if is_dark else "#D5D5D5")])

        self.tree.tag_configure("mapped", background="#1E3A29" if is_dark else "#EBFBEE",
                                foreground="#E2F0D9" if is_dark else "#1E4620")
        self.tree.tag_configure("unmapped", background="#4A1C1C" if is_dark else "#FFF5F5",
                                foreground="#FCE4EC" if is_dark else "#7A1C1C")
        self.tree.tag_configure("missing", background="#4A3B12" if is_dark else "#FFF8E1",
                                foreground="#FFE9A8" if is_dark else "#6B4E00")

    def update_data(self, records, warnings, has_mapping=False):
        """Shows records and warnings from core.excel_writer.build_records."""
        self._rows = []
        for r in records:
            metrics = sum(1 for m in METRIC_COLUMNS if r.get(m) is not None)
            mapped = r["Genotype"] != UNMAPPED and r["Sex"] != UNMAPPED
            if metrics < len(METRIC_COLUMNS):
                tag = "missing"
            else:
                tag = "mapped" if mapped else ("unmapped" if has_mapping else "")
            self._rows.append(({
                "table_id": r["Table ID"], "animal_id": r["Animal ID"], "timepoint": r["Timepoint"],
                "week": r["Week"], "genotype": r["Genotype"], "sex": r["Sex"], "runs": r["Runs"],
                "metrics": metrics,
            }, tag, mapped))

        self.only_unmapped.configure(state="normal" if has_mapping else "disabled")
        if not has_mapping:
            self.only_unmapped_var.set(False)
        self._set_warnings(warnings)
        self._refresh()

    def _set_warnings(self, warnings):
        self.view_switch.configure(values=[TABLES_VIEW, f"Warnings ({len(warnings)})"])
        self.view_switch.set(TABLES_VIEW)
        self._show_view(TABLES_VIEW)
        self.warnings_box.configure(state="normal")
        self.warnings_box.delete("1.0", "end")
        self.warnings_box.insert("1.0", "\n".join(f"• {w}" for w in warnings) or "No warnings.")
        self.warnings_box.configure(state="disabled")

    def _show_view(self, view):
        if view == TABLES_VIEW:
            self.warnings_box.pack_forget()
            self.table_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        else:
            self.table_container.pack_forget()
            self.warnings_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _sort_by(self, key):
        self._sort_reverse = not self._sort_reverse if key == self._sort_key else False
        self._sort_key = key
        self._refresh()

    def _refresh(self):
        """Redraws the table with the current sort column and unmapped filter."""
        def sort_value(row):
            value = row[0][self._sort_key]
            if self._sort_key == "timepoint":
                return sort_tags(value)
            return value if isinstance(value, int) else str(value).lower()

        rows = [r for r in self._rows if not (self.only_unmapped_var.get() and r[2])]
        self.tree.delete(*self.tree.get_children())
        for values, tag, _ in sorted(rows, key=sort_value, reverse=self._sort_reverse):
            shown = [values[k] for k, *_ in self.COLUMNS]
            shown[-1] = f"{values['metrics']}/{len(METRIC_COLUMNS)}"
            self.tree.insert("", "end", values=shown, tags=(tag,))

        arrow = " ▼" if self._sort_reverse else " ▲"
        for key, heading, *_ in self.COLUMNS:
            self.tree.heading(key, text=heading + (arrow if key == self._sort_key else ""))
