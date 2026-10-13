import os
import re
from collections import defaultdict
from statistics import fmean

from openpyxl import load_workbook

# Default target metric labels in Mouseframe Excel tables
DEFAULT_RAW_STRINGS = [
    "Stride length left front average in cm:",
    "Stride length right front average in cm:",
    "Stride length left hind average in cm:",
    "Stride length right hind average in cm:",
    "Overlap left average in cm:",
    "Overlap Right average in cm:",
    "Stride Width Front(L) average in cm:",
    "Stride Width Front(R) average in cm:",
    "Stride Width Hind(L) average in cm:",
    "Stride Width Hind(R) average in cm:"
]

# Rename mapping for unified metric labels
DEFAULT_RENAME_MAP = {
    "Stride Width Front(L) average in cm:": "Stride Width Front average in cm:",
    "Stride Width Front(R) average in cm:": "Stride Width Front average in cm:",
    "Stride Width Hind(L) average in cm:": "Stride Width Hind average in cm:",
    "Stride Width Hind(R) average in cm:": "Stride Width Hind average in cm:"
}

# Output metric columns in a stable order (after renaming)
METRIC_COLUMNS = list(dict.fromkeys(DEFAULT_RENAME_MAP.get(s, s) for s in DEFAULT_RAW_STRINGS))

SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm")

# Keys added by the parser that are not metrics
META_KEYS = ("Table ID", "Sheet_Source", "Runs", "Converted")

_SESSION_RE = re.compile(r"T\d+", re.IGNORECASE)
_WEEK_RE = re.compile(r"(\d+)(?:\+(\d+))?W(?:\+(\d+))?", re.IGNORECASE)
_SPLIT_RE = re.compile(r"[_\s\-]+")


def _to_number(value):
    """Returns value as float, accepting numeric strings with '.' or ',' decimals. None if not numeric."""
    if isinstance(value, bool):
        return None
    try:
        return float(str(value).strip().replace(",", "."))
    except ValueError:
        return None


def find_table_data(workbook_path, threshold=500.0):
    """
    Extracts table data from a raw Mouseframe Excel workbook.

    Every 'Table ID:' cell starts a table. The metric labels are read from the same column
    and their value from the column to the right, until the next 'Table ID' in that column.
    A table may contain several runs (separated by BREAK rows); each metric is averaged
    over the runs and the number of runs is stored in 'Runs'.
    """
    if not workbook_path.lower().endswith(SUPPORTED_EXTENSIONS):
        raise ValueError(
            f"Unsupported file type: {os.path.basename(workbook_path)}\n"
            "Only .xlsx / .xlsm workbooks are supported. Re-save .xls files as .xlsx in Excel."
        )

    raw_set = set(DEFAULT_RAW_STRINGS)
    wb = load_workbook(workbook_path, data_only=True)
    result_data = []

    try:
        for ws in wb.worksheets:
            max_row = ws.max_row
            for row in ws.iter_rows():
                for cell in row:
                    if not (isinstance(cell.value, str) and cell.value.strip().startswith("Table ID")):
                        continue

                    raw_table_id = re.sub(r"^Table ID\s*:?", "", cell.value.strip()).strip()
                    col_index = cell.column
                    table_values = defaultdict(list)
                    converted = 0

                    for r in range(cell.row + 1, max_row + 1):
                        label = ws.cell(row=r, column=col_index).value
                        if isinstance(label, str):
                            label = label.strip()
                            if label.startswith("Table ID"):
                                break  # next table stacked below in the same column
                        if label not in raw_set:
                            continue

                        value = _to_number(ws.cell(row=r, column=col_index + 1).value)
                        if value is None:
                            continue
                        # ponytail: unit heuristic, values above threshold assumed to be in 1/1000 of the expected unit
                        if threshold is not None and value > threshold:
                            value /= 1000.0
                            converted += 1
                        table_values[DEFAULT_RENAME_MAP.get(label, label)].append(value)

                    table_data = {
                        "Table ID": raw_table_id,
                        "Sheet_Source": ws.title,
                        "Runs": max((len(v) for v in table_values.values()), default=0),
                        "Converted": converted,
                    }
                    for key, values in table_values.items():
                        table_data[key] = fmean(values)
                    result_data.append(table_data)
    finally:
        wb.close()

    return result_data


def normalize_week(token):
    """'56+1W' / '56W+1' / '4w' -> '56W+1' / '4W'. None if token is not a week label."""
    m = _WEEK_RE.fullmatch(token.strip())
    if not m:
        return None
    week, plus_a, plus_b = m.groups()
    plus = plus_a or plus_b
    return f"{week}W+{plus}" if plus else f"{week}W"


def parse_table_id(table_id, sheet_source="Unknown"):
    """
    Splits a Table ID such as '2199_1R_T1_4W' into its parts.

    Returns (animal_id, timepoint, week):
      animal_id: Table ID without the session and week tokens ('2199_1R')
      timepoint: session token ('T1'), else week label, else the sheet name
      week:      normalized week label ('4W') or '' if absent
    """
    parts = [p for p in _SPLIT_RE.split(str(table_id).strip()) if p]
    session = week = None
    animal_parts = []
    for part in parts:
        if session is None and _SESSION_RE.fullmatch(part):
            session = part.upper()
        elif week is None and normalize_week(part):
            week = normalize_week(part)
        else:
            animal_parts.append(part)

    animal_id = "_".join(animal_parts) or str(table_id).strip()
    timepoint = session or week or sheet_source or "Unknown"
    return animal_id, timepoint, week or ""


def group_data_by_tag(data):
    """Groups extracted table entries by timepoint (session token, week label or sheet name)."""
    w_groups = defaultdict(list)
    for entry in data:
        _, timepoint, _ = parse_table_id(entry.get("Table ID", ""), entry.get("Sheet_Source", "Unknown"))
        w_groups[timepoint].append(entry)
    return w_groups


def sort_tags(tag):
    """Sort key for timepoint tags: numeric part first, then text ('T2' < 'T10', '56W' < '56W+1')."""
    nums = [int(n) for n in re.findall(r"\d+", str(tag))]
    return (0, nums, str(tag)) if nums else (1, [], str(tag))
