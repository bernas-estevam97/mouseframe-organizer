import customtkinter as ctk

from core import APP_NAME
from core.excel_writer import MODE_PASTEL_ACCENTS, MODE_COLUMNS_ONLY, MODE_SOFT_ROW_FILLS

HEADING_COLOR = ("#1F6AA5", "#5DADE2")

MAPPING_EXAMPLE = (
    "ID        Genotype  Sex\n"
    "2199_1R   WT        M\n"
    "3879      SCA2      F"
)


class TutorialFrame(ctk.CTkScrollableFrame):
    """In-app user guide explaining every input, option and output sheet."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._build_content()

    def _build_content(self):
        ctk.CTkLabel(self, text=f"{APP_NAME}: user guide", font=ctk.CTkFont(size=20, weight="bold")).pack(
            anchor="w", padx=15, pady=(15, 10))

        intro = (
            "This app reads the gait-analysis workbook exported by MouseFrame footprint tracking and writes one "
            "organized Excel file: one sheet per timepoint, a mean per animal, group statistics and a log of "
            "every data problem found. Typical use: pick the files, click Preview, check the Warnings view, "
            "then click Export workbook."
        )
        ctk.CTkLabel(self, text=intro, wraplength=820, justify="left").pack(anchor="w", padx=15, pady=(0, 15))

        self._add_section("1. Files", [
            ("MouseFrame workbook (.xlsx / .xlsm)",
             "The raw workbook from MouseFrame. Every sheet is scanned for cells starting with 'Table ID:'. "
             "Below each one the app reads the 'average in cm' rows (stride length, overlap, stride width). "
             "Old .xls files must be re-saved as .xlsx in Excel."),
            ("Mapping file (optional, CSV or Excel)",
             "A header row and 3 columns: ID, Genotype, Sex. Example:", MAPPING_EXAMPLE),
            ("How IDs are matched",
             "A mapping ID matches a Table ID when it equals its first parts (split on '_', '-' or spaces), "
             "ignoring case. '3879' matches '3879_T3_12W'; '2199_1R' matches '2199_1R_T1_4W' but not "
             "'2199_1L_T1_4W'. When several IDs match, the longest one wins. '219' never matches '2199'. "
             "An ID listed twice with different Genotype/Sex stops processing with an error."),
            ("Output folder (optional)",
             "Where the new workbook is saved. Empty: the folder of the MouseFrame workbook. The file name "
             "contains the date and time, so earlier results are never overwritten."),
            ("Remembered paths",
             "The last used files, folder and theme are restored when the app starts again."),
        ])

        self._add_section("2. Options", [
            (f"Group colors: {MODE_PASTEL_ACCENTS} (default)",
             "Pastel color on the Genotype and Group cells only; numbers stay on a white background."),
            (f"Group colors: {MODE_COLUMNS_ONLY}",
             "Genotype, Sex and Group columns without any fill. Best for pasting into GraphPad Prism, R or SPSS."),
            (f"Group colors: {MODE_SOFT_ROW_FILLS}",
             "Pastel fill on the whole row, chosen by genotype and sex. Colors need a mapping file."),
            ("Decimal-fix threshold (default 500)",
             "When numbers are pasted into Excel set to a language that uses ',' as the decimal mark, a value "
             "such as 2.456 cm can be read as 2456. Gait values in cm are far below 500, so any value (positive "
             "or negative) above the threshold is divided by 1000 and listed in the warnings. A damaged value "
             "that stays below the threshold (for example 0.312 read as 312 with threshold 300) is not fixed, "
             "so keep the threshold above your largest real value but below the smallest damaged one."),
            ("Add summary sheets",
             "'Summary_Per_Animal': one row per animal per timepoint with the mean of every metric. If an animal "
             "was recorded more than once at the same timepoint, its tables are averaged here. With a mapping "
             "file, 'Group_Statistics' adds N animals, Mean and SEM per Timepoint × Genotype × Sex. "
             "Animals, not tables, are the statistical unit."),
        ])

        self._add_section("3. How tables are read", [
            ("Timepoints, Animal IDs and weeks",
             "A Table ID such as '2199_1R_T1_4W' is split into Animal ID ('2199_1R'), session ('T1') and week "
             "('4W'). Sheets are grouped by session, so 56W, 56W+1 and 56W+2 tables of session T14 all go to "
             "sheet T14; the Week column shows the variant. IDs without a session use the week, then the sheet name."),
            ("Runs",
             "A table may hold several runs separated by BREAK rows. Metrics are averaged over the runs and "
             "the Runs column shows how many there were."),
        ])

        self._add_section("4. Preview and export", [
            ("Preview",
             "Reads the workbook without writing anything. Green rows are matched to the mapping file, red rows "
             "are not, amber rows miss one or more metrics (Metrics shows for example 6/8). Click a column header "
             "to sort; tick 'Only unmapped' to list the IDs missing from the mapping file. The Warnings view lists "
             "every data problem."),
            ("Export workbook",
             "Writes the organized workbook. When it is done you can open it right away, or later with the "
             "Open file and Open folder buttons at the bottom of the window."),
            ("Output sheets",
             "Summary_Per_Animal, Group_Statistics, one sheet per timepoint (T1, T2, …), Processing_Log "
             "(warnings: duplicated Table IDs, missing metrics, decimal fixes, Table IDs in the wrong week sheet, "
             "unmapped tables, unused mapping IDs) and Info (app version, date, input files and settings used)."),
            ("When something goes wrong",
             "Error messages say what to fix. Technical details are saved in the file "
             "mouseframe_organizer.log in your user folder; send it along when reporting a problem."),
        ])

    def _add_section(self, title, items):
        """A card with a title and (heading, text) or (heading, text, monospace example) items."""
        card = ctk.CTkFrame(self, corner_radius=8)
        card.pack(fill="x", padx=15, pady=8)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=15, pady=(12, 6))

        for heading, body, *example in items:
            ctk.CTkLabel(card, text=heading, font=ctk.CTkFont(weight="bold"), text_color=HEADING_COLOR).pack(
                anchor="w", padx=20, pady=(4, 1))
            ctk.CTkLabel(card, text=body, wraplength=780, justify="left").pack(anchor="w", padx=20, pady=(0, 6))
            if example:
                ctk.CTkLabel(card, text=example[0], justify="left", font=ctk.CTkFont(family="Consolas", size=13),
                             fg_color=("gray90", "gray20"), corner_radius=6).pack(
                    anchor="w", padx=20, pady=(0, 8), ipadx=10, ipady=6)
