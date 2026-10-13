import customtkinter as ctk
from tkinter import ttk


class PreviewTableFrame(ctk.CTkFrame):
    """
    Data preview widget using tkinter.ttk Treeview embedded in a CustomTkinter frame.
    Displays parsed Table IDs, assigned Time Tags, matched Genotype, and Sex.
    Supports Light and Dark appearance modes.
    """

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
        columns = ("table_id", "timepoint", "genotype", "sex", "group", "metrics_count")
        
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.tree = ttk.Treeview(
            self.table_container,
            columns=columns,
            show="headings",
            height=8
        )

        # Configure headings
        self.tree.heading("table_id", text="Table ID")
        self.tree.heading("timepoint", text="Time Tag")
        self.tree.heading("genotype", text="Genotype")
        self.tree.heading("sex", text="Sex")
        self.tree.heading("group", text="Group")
        self.tree.heading("metrics_count", text="Metrics Parsed")

        # Column widths
        self.tree.column("table_id", width=160, anchor="w")
        self.tree.column("timepoint", width=90, anchor="center")
        self.tree.column("genotype", width=110, anchor="center")
        self.tree.column("sex", width=80, anchor="center")
        self.tree.column("group", width=130, anchor="center")
        self.tree.column("metrics_count", width=110, anchor="center")

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

    def update_data(self, data, mapping_dict=None):
        """Populates the preview table with parsed records."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not data:
            return

        from core.parser import group_data_by_tag
        from core.mapping import match_table_id

        w_groups = group_data_by_tag(data)

        for tag, group_data in w_groups.items():
            for entry in group_data:
                table_id = entry.get("Table ID", "")
                metric_keys = [k for k in entry.keys() if k not in ("Table ID", "Sheet_Source")]

                geno, sex = None, None
                if mapping_dict:
                    geno, sex = match_table_id(table_id, mapping_dict)

                geno_str = geno if geno else "Unmapped"
                sex_str = sex if sex else "Unmapped"
                group_str = f"{geno_str} {sex_str}".strip()

                tag_name = "mapped" if (geno and sex) else ("unmapped" if mapping_dict else "")

                self.tree.insert(
                    "",
                    "end",
                    values=(table_id, tag, geno_str, sex_str, group_str, len(metric_keys)),
                    tags=(tag_name,)
                )
