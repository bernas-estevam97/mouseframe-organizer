# MouseFrame Organizer

Desktop app (Python + CustomTkinter) that turns the raw gait-analysis workbook exported from **MouseFrame** footprint tracking into a clean, analysis-ready Excel file. It organizes every table by **timepoint**, attaches **genotype** and **sex** from an optional mapping file, computes the **mean per animal**, and reports **N / Mean / SEM per group**.

## Quick start

```bash
python -m venv venv
venv\Scripts\activate          # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
python main.py
```

Requires Python 3.10+ (tested with Python 3.12, pandas 3.0.6, openpyxl 3.1.5, customtkinter 6.0).

Run the tests with:

```bash
python -m unittest discover tests
```

## Input: raw MouseFrame workbook (.xlsx / .xlsm)

The example `data sheet classic FP SCA2.xlsx` has one sheet per timepoint (`4W`, `8W`, …, `56W+1`, `84W+2`). Inside each sheet, tables sit side by side, starting every 7 columns in row 1:

```text
A1  Table ID: 2199_1R_T1_4W
A2  Entry Number | Points | Measurement type | Value(px) | Value(cm)
A3  1            | Red L1 → Red L2 | Stride length left front | 231.262 | 8.559
...
A8  Stride length left front average in cm:   | 8.152      <- the value the app reads
...
A60 ------------------------------BREAK------------------------------
... (optional further runs of the same animal)
```

For every `Table ID:` cell the app reads, from the same column downwards (until the next `Table ID` in that column), the rows whose label is one of:

| Label in MouseFrame | Output column |
| --- | --- |
| Stride length left/right front average in cm: | same name |
| Stride length left/right hind average in cm: | same name |
| Overlap left / Overlap Right average in cm: | same name |
| Stride Width Front(L) / Front(R) average in cm: | Stride Width Front average in cm: |
| Stride Width Hind(L) / Hind(R) average in cm: | Stride Width Hind average in cm: |

The value is taken from the cell to the right of the label. Numeric text with a comma decimal (`"7,5"`) is accepted. When a table contains several runs (separated by `BREAK` rows), each metric is averaged across runs and the number of runs is stored in the **Runs** column.

`.xls` files are rejected with a clear message: re-save them as `.xlsx`.

### Table ID convention

A Table ID such as `2199_1R_T1_4W` is split on `_`, `-` and spaces:

| Part | Example | Meaning |
| --- | --- | --- |
| Animal ID | `2199_1R` | everything except the session and week tokens |
| Timepoint | `T1` | session token `T<number>`; used to group sheets |
| Week | `4W` | `4W`, `56+1W`, `68W+1` (normalized to `56W+1`, `68W+1`) |

If a Table ID has no `T<n>` token, the week label is the timepoint; if it has neither, the sheet name is used. `56W`, `56W+1`, `56W+2`, `56W+3`, … are the same timepoint by design: they share session `T14` and land on the same `T14` sheet and in the same `Group_Statistics` rows. The **Week** column only records which variant each table came from.

## Input: mapping file (optional, .csv / .xlsx)

A header row plus three columns: subject ID, genotype, sex. Extra columns are ignored. CSV separator (`,` or `;`) is detected automatically.

```text
ID,Genotype,Sex
2199_1R,WT,M
3879,SCA2,F
```

Matching rules:

- A mapping ID matches when its tokens are the leading tokens of the Table ID, ignoring case: `3879` matches `3879_T3_12W`; `2199_1R` matches `2199_1R_T1_4W` but not `2199_1L_T1_4W`.
- The longest matching ID wins (`2199_1R` beats `2199`).
- Partial numbers never match (`219` does not match `2199`).
- The same ID listed twice with different genotype/sex stops processing with an error.
- IDs are read as text, so leading zeros and blank cells cannot corrupt them.

Sex labels `M/Male/Macho/Masculino` and `F/Female/Fêmea/Feminino` get the male/female color palettes; anything else gets a neutral palette.

## Output workbook

Saved as `MFrame_clean_output_<YYYY-MM-DD_HH-MM-SS>.xlsx` in the chosen output folder (default: next to the input file). The file is written to a temporary `.part` file first and renamed on success, so an interrupted run never leaves a corrupt workbook.

| Sheet | Content |
| --- | --- |
| `Summary_Per_Animal` | One row per animal per timepoint: Animal ID, Timepoint, number of tables, Table ID, Week, Genotype, Sex, Group, and the mean of each metric. Repeated tables of one animal at one timepoint are averaged here. |
| `Group_Statistics` | Only with a mapping file. Per Timepoint × Genotype × Sex: N animals, then Mean and SEM of every metric. Animals (not tables) are the statistical unit. SEM is empty when N = 1. |
| `T1`, `T2`, … | One sheet per timepoint with every table: Table ID, Animal ID, Week, Genotype, Sex, Group, Runs, Source Sheet, metrics. |
| `Processing_Log` | Only when there are warnings (see below). |

The two summary sheets are produced when **Generate Summary Sheets** is checked. Unmapped tables show `Unmapped` in Genotype/Sex/Group.

### Visual styles

| Style | Effect |
| --- | --- |
| Pastel Header Tag + Explicit Columns | Pastel fill on Genotype and Group cells only. |
| Explicit Columns Only (Clean White Grid) | No fills; best for pasting into Prism, R or SPSS. |
| Soft Pastel Row Fills | Pastel fill on the whole row. |

Colors need a mapping file. Each genotype gets a palette index; sex chooses the palette.

### Unit conversion threshold

Any metric value above the threshold (default `500`) is divided by 1000, on the assumption that it was exported in a smaller unit. Every conversion is reported in `Processing_Log`. The real example file has a maximum value of 11.2 cm, so no conversion happens there.

### Data-quality warnings

Shown in the status bar after **Preview** (first 5) and written to `Processing_Log` after export:

- Same Table ID more than once in one sheet (16W and 20W in the example file).
- Table without any metric, or missing some of the expected metrics.
- Values divided by 1000 by the threshold rule.
- Table ID week that does not match the week sheet it is in.
- Tables without a mapping match, and mapping IDs that matched no table.

## Using the app

1. **Raw Mouseframe Excel File**: pick the `.xlsx` export. Paths pasted with quotes ("Copy as path") are accepted.
2. **Metadata Mapping File** (optional): pick the CSV/Excel mapping. **Clear** removes it.
3. **Output Directory** (optional): must exist; otherwise processing stops with an error.
4. Choose a visual style and the threshold.
5. **🔍 Preview Data**: parses without writing anything. Green rows are mapped, red rows are unmapped.
6. **🚀 Process & Export Excel**: writes the workbook and offers to open its folder.

Both buttons are disabled while a job runs. Errors are shown in a dialog and logged with full traceback to `~/mouseframe_organizer.log`.

## Project layout

```text
main.py               entry point: logging, Windows taskbar ID, icon (PyInstaller-aware)
core/parser.py        reads tables from the raw workbook; Table ID parsing; timepoint sort
core/mapping.py       loads the mapping file; Table ID -> (genotype, sex) matching
core/excel_writer.py  records + warnings, summaries, group statistics, styled workbook export
gui/app.py            main window; background jobs; thread-safe UI updates via a queue
gui/preview_table.py  preview Treeview (light/dark aware)
gui/tutorial.py       in-app user guide
tests/test_core.py    unittest suite on a synthetic workbook
CLAUDE.md             project guide: commands, conventions, domain rules, change history
```

## Building a Windows executable

Recommended: **one-folder build, shipped as a zip**.

```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --onedir --name MouseFrameOrganizer --icon assets/app_icon.ico --add-data "assets;assets" main.py
```

Result: `dist/MouseFrameOrganizer/` (about 80 MB, about 30 MB zipped). Zip the folder and share it. Users unzip it and run `MouseFrameOrganizer.exe`. The exe must stay next to its `_internal` folder.

Why one-folder and not `--onefile` (measured on this project):

| | `--onedir` | `--onefile` |
| --- | --- | --- |
| Start-up to window | ~1.0 s | ~2.7 s (unpacks about 80 MB to `%TEMP%` on every launch) |
| Size | 80 MB folder | 33 MB single exe |
| Antivirus false positives | rare | more common (self-extracting exe) |
| Leftovers | none | `%TEMP%\_MEI*` folders stay behind if the app is killed |

`--onefile` still works. Use it if handing out a single file matters more than start-up time:

```bash
pyinstaller --noconfirm --windowed --onefile --name MouseFrameOrganizer --icon assets/app_icon.ico --add-data "assets;assets" main.py
```

Notes:

- Build on Windows, inside the project venv, so only the packages in `requirements.txt` get bundled. PyInstaller cannot cross-compile.
- No `--add-data` is needed for customtkinter: PyInstaller's bundled hook copies its themes and fonts.
- `main.py` replaces the missing `sys.stdout`/`sys.stderr` of `--windowed` builds. Without that, a customtkinter warning would crash the app on start-up, for example when the unzipped folder sits in a path longer than 260 characters.
- Errors from the built app go to `~/mouseframe_organizer.log`.
