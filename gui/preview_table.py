import customtkinter as ctk
from tkinter import ttk

from core.parser import sort_tags
from core.excel_writer import LEAD_COLUMNS, UNMAPPED


class PreviewTableFrame(ctk.CTkFrame):
    """
    Data preview widget using tkinter.ttk Treeview embedded in a CustomTkinter frame.
    Displays parsed Table IDs, Animal IDs, timepoints, matched Genotype/Sex and run counts.
    Supports Light and Dark appearance modes.
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
        ("metrics_count", "Metrics", 60, "center"),
    ]

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.title_label = ctk.CTkLabel(
            self,
            text="Data Preview & Metadata Validation",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.title_label.pack(anchor="w", padx=10, pady=(10, 5))

        # Setup container for Treeview and scrollbars
        self.table_container = ctk.CTkFrame(self)
        self.table_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Treeview widget
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.tree = ttk.Treeview(
            self.table_container,
            columns=[c[0] for c in self.COLUMNS],
            show="headings",
            height=10
        )
        for key, heading, width, anchor in self.COLUMNS:
            self.tree.heading(key, text=heading)
            self.tree.column(key, width=width, anchor=anchor)

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(self.table_container, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(self.table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.apply_theme_colors(ctk.get_appearance_mode())

    def apply_theme_colors(self, mode="Dark"):
        """Configures Treeview colors depending on Light/Dark appearance mode."""
        is_dark = mode.lower() == "dark"

        bg_color = "#2A2D2E" if is_dark else "#FFFFFF"
        fg_color = "#FFFFFF" if is_dark else "#000000"
        field_bg = "#2A2D2E" if is_dark else "#FFFFFF"
        header_bg = "#1F1F1F" if is_dark else "#E0E0E0"
        header_fg = "#FFFFFF" if is_dark else "#000000"

        self.style.configure(
            "Treeview",
            background=bg_color,
            foreground=fg_color,
            fieldbackground=field_bg,
            rowheight=24,
            font=("Segoe UI", 10)
        )
        self.style.configure(
            "Treeview.Heading",
            background=header_bg,
            foreground=header_fg,
            font=("Segoe UI", 10, "bold")
        )
        self.style.map("Treeview.Heading", background=[('active', '#3B8ED0' if is_dark else '#B0C4DE')])

        # Row tag colors for mapped vs unmapped records
        mapped_bg = "#1E3A29" if is_dark else "#EBFBEE"
        mapped_fg = "#E2F0D9" if is_dark else "#1E4620"
        unmapped_bg = "#4A1C1C" if is_dark else "#FFF5F5"
        unmapped_fg = "#FCE4EC" if is_dark else "#7A1C1C"

        self.tree.tag_configure("mapped", background=mapped_bg, foreground=mapped_fg)
        self.tree.tag_configure("unmapped", background=unmapped_bg, foreground=unmapped_fg)

    def update_data(self, records, has_mapping=False):
        """Populates the preview table with records from core.excel_writer.build_records."""
        self.tree.delete(*self.tree.get_children())

        for r in sorted(records, key=lambda r: sort_tags(r["Timepoint"])):
            metrics = sum(1 for k, v in r.items() if k not in LEAD_COLUMNS and v is not None)
            mapped = r["Genotype"] != UNMAPPED and r["Sex"] != UNMAPPED
            tag_name = "mapped" if mapped else ("unmapped" if has_mapping else "")
            self.tree.insert(
                "",
                "end",
                values=(r["Table ID"], r["Animal ID"], r["Timepoint"], r["Week"],
                        r["Genotype"], r["Sex"], r["Runs"], metrics),
                tags=(tag_name,)
            )
