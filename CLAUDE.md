# CLAUDE.md: MouseFrame Organizer

Last updated: 2026-10-20

Project guide, loaded automatically by Claude Code. It is the guide for humans and agents: what the project is, how to run and test it, the rules the code follows, and the history of what changed and why. Update it (and the date above) whenever you change behavior, commands or conventions.

## Project at a glance

- Windows desktop app (Python + CustomTkinter). Turns raw MouseFrame gait-analysis workbooks (`.xlsx`) into an organized Excel workbook: one sheet per timepoint, per-animal means, group N/Mean/SEM, and a warnings log.
- Users: lab researchers (non-programmers). They run a built `.exe`; error messages must be plain and actionable.
- User-facing documentation: `README.md` (input format, output sheets, matching rules, build). In-app guide: `gui/tutorial.py`. Keep all three consistent when behavior changes.
- Git: single branch, one initial commit (`3ae43b7`). Commit only when the user asks.

## Development commands

Run from the project root. Shell examples use Git Bash; PowerShell equivalents use `venv\Scripts\python.exe`.

| Task | Command |
| --- | --- |
| Create venv (Python 3.12) | `python -m venv venv` |
| Install dependencies | `venv/Scripts/python.exe -m pip install -r requirements.txt` |
| Install build tool | `venv/Scripts/python.exe -m pip install pyinstaller` |
| Run the app | `venv/Scripts/python.exe main.py` |
| Run tests | `venv/Scripts/python.exe -m unittest discover tests` |
| Check for outdated packages | `venv/Scripts/python.exe -m pip list --outdated` |
| Build exe (recommended) | `venv/Scripts/pyinstaller --noconfirm --windowed --onedir --name MouseFrameOrganizer --icon assets/app_icon.ico --add-data "assets;assets" main.py` |
| Build single exe (alternative) | same command with `--onefile` instead of `--onedir` |
| Run pipeline without GUI | `venv/Scripts/python.exe -c "from core.excel_writer import process_and_export_file as p; print(p('in.xlsx', 'map.csv', 'out_dir'))"` |

- Build output goes to `dist/MouseFrameOrganizer/`; ship that folder zipped. `build/`, `dist/` and `*.spec` are git-ignored. A local `main.spec` from an earlier `--onefile` build named `main` may exist; it is stale, use the command above instead.
- Runtime log of the app (including built exe): `~/mouseframe_organizer.log`. Last paths and theme: `~/mouseframe_organizer_settings.json` (written on window close).
- App name and version: `core/__init__.py` (`APP_NAME`, `__version__`, now `1.1.0`). Bump the version when behavior changes; it appears in the window title, header and the `Info` sheet.
- Pinned dependencies (`requirements.txt`): `customtkinter==6.0.0`, `openpyxl==3.1.5`, `pandas==3.0.6`. PyInstaller is a build tool only and is not listed there.
- When reading untrusted workbooks from a script, use `python -I` (isolated mode).

## Architecture and data flow

```text
main.py                 logging -> stdout/stderr guard -> MouseframeOrganizerApp
gui/app.py              header, form + validation (_read_inputs), fixed footer (progress, status, Open file/folder)
                        -> _run_in_background(work, on_success, error_title, indeterminate)
  Preview  -> core.parser.find_table_data -> core.mapping.load_mapping_file -> core.excel_writer.build_records
           -> gui.preview_table (sortable table, "Only unmapped" filter, Warnings view)
  Export   -> core.excel_writer.process_and_export_file
core/parser.py          workbook -> list of table dicts {Table ID, Sheet_Source, Runs, Converted, <metrics>}
                        parse_table_id(): Table ID -> (Animal ID, Timepoint, Week)
core/mapping.py         mapping file -> {ID: (genotype, sex)}, genotype colour index; match_key()
core/excel_writer.py    build_records() -> records + warnings (shared by preview and export)
                        build_summary() / build_group_stats() -> DataFrames
                        _write_df() / _apply_group_fills() -> styled sheets; atomic save via .part file
```

- `core/` has no GUI imports and can be used from scripts and tests. `gui/` holds all Tk code.
- `process_and_export_file(...)` returns `(output_path, warnings)`. `load_mapping_file(...)` returns `(mapping_dict, genotype_index)`; `({}, {})` when no path is given.
- Output sheet order: `Summary_Per_Animal`, `Group_Statistics` (only with a mapping), `T1`…`Tn` (sorted by `sort_tags`), `Processing_Log` (only with warnings), `Info` (always: version, date, input paths, settings).
- Organizer tab is not scrollable on purpose (a scrollable frame around it hid the footer below the fold and nested two scroll areas). The minimum window size (900×740) keeps the form and a few preview rows visible.

## Domain rules (do not break)

- **Timepoint = session token `T<n>` in the Table ID.** `56W`, `56W+1`, `56W+2`, `56W+3`, … all belong to the same timepoint (`T14`), confirmed by the user. Never split timepoints by week suffix; `Week` only records the variant. Fallbacks: week token, then sheet name.
- **Animal ID** = Table ID without the session and week tokens (`2199_1R_T1_4W` -> `2199_1R`). `1777_1R` and `1777_2R` are different animals.
- **Metrics** are the "average in cm" rows listed in `DEFAULT_RAW_STRINGS`; value is in the cell right of the label. Stride Width `(L)`/`(R)` merge into one Front and one Hind column (`DEFAULT_RENAME_MAP`). Output order is `METRIC_COLUMNS`.
- **Runs**: a table can hold several runs (separated by `BREAK` rows); metrics are averaged over runs, `Runs` stores the count. A table ends at the next `Table ID` in the same column.
- **Statistics unit is the animal**: `Summary_Per_Animal` averages repeated tables per Animal ID + Timepoint first; `Group_Statistics` (N, Mean, SEM with ddof=1) is computed over those per-animal rows.
- **Mapping match**: mapping ID tokens must equal the leading tokens of the Table ID (case-insensitive, split on `_`, `-`, spaces); longest key wins. Never substring matching. IDs are read as text (`dtype=str`). Conflicting duplicate IDs raise `ValueError`.
- **Decimal-fix threshold**: values whose absolute value is above the threshold (default 500) are divided by 1000 and counted as a warning. Reason (confirmed by the user, 2026-10-20): in Excel set to a `,`-decimal language, a pasted `2.456` cm is read with `.` as thousands separator and becomes `2456`. Negative values (Overlap) are fixed too. Damaged values that stay below the threshold (originals < threshold/1000 cm) are missed; marked `ponytail:` in `parser.py`.
- **Never fail silently**: data problems become entries in the warnings list (Warnings view + `Processing_Log`); bad inputs (missing file, missing folder, `.xls`, bad mapping) raise clear exceptions shown in a dialog.

## Code style and conventions

- Python 3.12, 4-space indent, f-strings. Double quotes in new code (some older lines use single quotes; do not churn them). No type hints (match existing code). No linter or formatter is configured; aim for ≤ 120 chars per line (some GUI widget lines in `gui/app.py` are longer and may stay).
- Names: `snake_case` functions/variables, `UPPER_CASE` module constants, `_leading_underscore` for module-private helpers, `CamelCase` classes. Regexes precompiled as `_NAME_RE` module constants.
- Docstrings: one or a few lines saying what the function returns or does. Comments only for the non-obvious "why".
- Minimalism ("ponytail"): standard library first, then existing dependencies; no new dependency for what a few lines do; no speculative parameters, layers or config. Deliberate shortcuts with a known ceiling get a `# ponytail: <ceiling>, <upgrade path>` comment.
- Shared logic lives once in `core/` (for example `build_records` serves both preview and export); do not duplicate it in the GUI.
- GUI threading: worker threads never touch Tk. Use `_run_in_background(work, on_success, error_title)`; from a worker, schedule UI work with `self._call_ui(fn)` (queue drained on the UI thread every 50 ms). Capture exception messages inside the `except` block before scheduling (the `as e` name is cleared after it).
- Errors: core raises built-in exceptions (`FileNotFoundError`, `NotADirectoryError`, `ValueError`) with user-readable messages; the GUI shows `str(e)` and logs the traceback with `log.exception`.
- Files written by the app: save to `<name>.part`, then `os.replace` (no half-written workbooks). Never write `NaN` into cells; write `None`.
- Excel styling constants (fonts, fills, borders, pastel palettes) live at the top of `core/excel_writer.py`; reuse them.
- `main.py` must replace `None` `sys.stdout`/`sys.stderr` before importing customtkinter (windowed builds crash otherwise).
- GUI look: no emoji in labels or buttons (they render tiny and monochrome in Tk on Windows). Export is the primary (filled) button, Preview the secondary (outlined) one. Hints use `HINT_COLOR` in `gui/app.py`. ttk widgets do not follow customtkinter scaling: scale pixel sizes with `ctk.ScalingTracker.get_widget_scaling`.
- UI text and code comments are in English. Sex labels accept English and Portuguese (`MALE_LABELS`, `FEMALE_LABELS`).
- Tests: stdlib `unittest` only, in `tests/test_core.py`, using a synthetic workbook built in a temp dir (no real data needed). Add one test per new parsing/matching rule or bug fix. Run the suite before claiming work is done.
- Docs: when behavior changes, update `README.md`, `gui/tutorial.py` and this file in the same change.

## Reference data

- `data sheet classic FP SCA2.xlsx` (example raw export; not committed, may be absent from the folder — ask the user for a copy if a real-data check is needed): 37 sheets (`4W` to `84W+2`), 600 tables, all `Table ID` cells in row 1, one table every 7 columns.
- 55 distinct animals. Table ID formats seen: `2199_1R_T1_4W`, `3879_T3_12W`, `3220_T14_56+1W`, `3220_T17_68W+1`.
- Sessions `T14` to `T21` each cover several week sheets (`56W`, `56W+1`, `56W+2`, …). Confirmed by the user (2026-10-13): `56W`, `56W+1`, `56W+2`, `56W+3`, … all belong to the same timepoint (`T14`). Grouping by the `T<n>` token is correct; `Week` only records the variant.
- Tables contain several runs separated by `BREAK` rows.
- Duplicated Table IDs: `2200_1R_T4_16W`, `2200_1L_T4_16W`, `2199_1R_T4_16W`, `2273_1R_T4_16W` (16W), `2221_1R_T5_20W` (20W).
- All 8 output metrics are present in every table. Maximum metric value is 11.188, so the threshold rule never triggers.

## Review findings (original code) and status

| # | Severity | Problem | Status |
| --- | --- | --- | --- |
| 1 | High | Error callbacks used the `except ... as e` variable after the block ended. Python deletes `e` at that point, so every error dialog would crash with `NameError` instead of showing the error. | Fixed: message captured before scheduling. |
| 2 | High | Mapping used substring matching with first-match-wins: key `219` matched `2199`, key `1` matched almost everything. Wrong genotype/sex was assigned silently. | Fixed: whole-token prefix match, longest key wins. |
| 3 | High | Mapping IDs read by pandas as numbers: a blank cell in the ID column turns `2199` into `2199.0`, which never matches. Blank genotype/sex became the text `nan`. | Fixed: read as text, blanks handled. |
| 4 | Low | Missing metric values became `NaN` and were written as numeric cells with an empty value node (`<v/>`). | Fixed: `NaN` written as a truly empty cell. |
| 5 | Medium | Worker threads called `self.after()` directly. Tkinter is not thread-safe: it usually works under `mainloop`, but raises `RuntimeError: main thread is not in main loop` when the loop is not running (seen in the scripted GUI test). | Fixed: queue drained on the UI thread. |
| 6 | Medium | `self.geometry("980, 820")` is an invalid geometry string. It was ignored and the window opened at minimum size. | Fixed: `"980x820"`. |
| 7 | Medium | A mapping path that did not exist was silently ignored: everything came out `Unmapped` without notice. | Fixed: error in GUI and core. |
| 8 | Medium | An output folder that did not exist was silently replaced by the input folder. | Fixed: error. |
| 9 | Medium | Duplicate Table IDs, missing metrics, unit conversions and unmapped IDs were never reported. | Fixed: warnings in status bar + `Processing_Log` sheet. |
| 10 | Medium | Summary checkbox promised "Mean ± SEM per Genotype/Sex", but only a per-animal mean was produced. | Fixed: new `Group_Statistics` sheet (N, Mean, SEM; animal = unit). |
| 11 | Medium | Table scan ran to the bottom of the sheet: a table stacked below another in the same column would leak its values into the first table. | Fixed: stop at the next `Table ID`. |
| 12 | Medium | File dialog offered `.xls`, which openpyxl cannot read (cryptic error). | Fixed: `.xlsx`/`.xlsm` only, clear message. |
| 13 | Medium | Workbook saved in place: closing the app during save left a corrupt file. | Fixed: save to `.part`, then atomic rename. |
| 14 | Low | Sheet titles were not sanitized. | Not needed: timepoint tags are `T<n>`/week regex matches or existing sheet names, so always valid. A sanitizer was added, then removed in the ponytail review. |
| 15 | Low | Numeric text with comma decimals (`"7,5"`) was skipped silently. | Fixed: parsed. |
| 16 | Low | Number of runs per table was hidden; runs were averaged without notice. | Fixed: `Runs` column. |
| 17 | Low | No Animal ID column: longitudinal analysis needed manual ID parsing. | Fixed: `Animal ID` + `Week` columns. |
| 18 | Low | Metric column order depended on which table came first. | Fixed: canonical order. |
| 19 | Low | Theme "System" always gave the light preview table, even in dark mode. | Fixed. |
| 20 | Low | Preview button stayed active during preview; double clicks started parallel jobs. | Fixed: both buttons disabled while busy. |
| 21 | Low | "Open folder" ran `open` on Linux (only valid on macOS). | Fixed: `xdg-open` on Linux. |
| 22 | Low | Icon path broke in a PyInstaller build; PNG icon reference could be garbage-collected. | Fixed. |
| 23 | Low | Sex label typo `fêmica`; Portuguese labels incomplete. | Fixed. |
| 24 | Low | Duplicated styling code (two copies), dead attributes (`_parsed_data_cache`, …), unused `numpy`. | Removed. |
| 25 | Low | No logging, no tests. | Added log file `~/mouseframe_organizer.log` and `tests/test_core.py`. |

## Changes by file

- `core/parser.py`: robust table scan, comma decimals, `Runs`/`Converted` counters, `.xls` rejection, `parse_table_id`, `normalize_week`, `METRIC_COLUMNS`, tuple-based `sort_tags`.
- `core/mapping.py`: text dtype, CSV separator sniffing, BOM-safe, conflict detection, `match_key` (token prefix, longest wins).
- `core/excel_writer.py`: `build_records` (shared by preview and export, collects warnings), `build_summary`, `build_group_stats`, `_write_df`, `_apply_group_fills`, atomic save. `process_and_export_file` now returns `(output_file, warnings)`.
- `gui/app.py`: input validation in one place, `_run_in_background`, UI-call queue, fixed geometry/theme/open-folder.
- `gui/preview_table.py`: shows Animal ID, Timepoint, Week, Runs; sortable; unmapped filter; Warnings view.
- `gui/tutorial.py`: matching rules, summary sheets, Runs, Processing_Log explained.
- `main.py`: file logging, Tk callback exception logging, PyInstaller-aware assets path, `None` stdout/stderr guard for windowed builds.
- `requirements.txt`: only direct dependencies, pinned: `customtkinter==6.0.0`, `openpyxl==3.1.5`, `pandas==3.0.6`. Dropped `numpy` (comes with pandas) and `pillow` (only used by a PNG-icon fallback for a file that does not exist; fallback removed).
- `README.md`: rewritten to describe the current behavior.
- `tests/test_core.py`: 8 unittest cases.

## Ponytail review cleanup (2026-10-13)

Removed: `safe_sheet_title`, unused `raw_strings`/`rename_map` parameters of `find_table_data`, the always-empty "extra metrics" branch, the `GROUP_KEYS` constant, and the third return value of `load_mapping_file` (now `(mapping_dict, genotype_index)`). Shortened `_to_number`, the UI-queue drain loop and logging setup (file only, no console handler).

## Executable build (2026-10-13)

- Strategy: PyInstaller `--windowed --onedir`, shipped as a zip. Command in README.
- Measured: onedir ~1.0 s to window, 80 MB folder; onefile ~2.7 s, 33 MB exe.
- Bug found while testing: in `--windowed` builds `sys.stderr` is `None`. customtkinter writes a warning to it when its font fails to load (seen with a path over 260 characters), and the app crashed on start-up. Fixed in `main.py` by replacing `None` streams with `os.devnull` before importing customtkinter. Verified: the same long-path build now starts.
- PyInstaller 6.22.3 is installed in the venv as a build tool. It is not in `requirements.txt`.

## GUI and usability review (2026-10-20)

Review of look and wording, then fixed:

- Tutorial said the threshold converted mm/µm to cm (wrong factor). Replaced by the real reason (decimal-separator damage), renamed "Decimal-fix threshold"; negative damaged values are now fixed too (`abs()`).
- Warnings were cut to 5 in the status label: now a Warnings view next to the preview table.
- Status and progress were below the fold inside a scrollable tab: now a fixed footer, with Open file / Open folder buttons for the last export. Preview uses an indeterminate progress bar.
- Error dialogs now name the log file.
- New `Info` sheet records version, date, input/mapping paths and settings.
- Look: emoji removed, primary/secondary buttons, aligned Browse buttons with a small ✕ to clear the mapping, hints in their own grid column, Treeview with themed `CTkScrollbar`, flat headers, DPI-scaled row height, selection color.
- Wording: style names "Color genotype cells" / "No colors" / "Color whole rows", "Add summary sheets", one app name "MouseFrame Organizer" with version.
- Preview: sortable columns, "Only unmapped" filter, metrics shown as `n/8` with amber rows when incomplete.
- Default theme "System"; last paths and theme remembered.
- Dead code removed: `group_data_by_tag`, `match_table_id` (tests use `match_key`).
- Considered and skipped: drag-and-drop (needs `tkinterdnd2`). Future idea from the user: try PySide6; `core/` has no GUI imports, so only `gui/` would be rewritten.

## Behavior changes to know

- `process_and_export_file` returns a tuple `(path, warnings)` instead of a path.
- Mapping IDs must now equal the leading token(s) of the Table ID. A mapping that relied on matching a fragment in the middle of the ID (for example `1R`) no longer matches.
- `Summary_Per_Animal` groups by Animal ID + Timepoint (was Table ID + Timepoint). Column order: Animal ID, Timepoint, Tables, Table ID, Week, Genotype, Sex, Group, metrics.
- Per-timepoint sheets have new columns `Animal ID`, `Week`, `Runs`, `Source Sheet`.
- Fully unmapped rows show Group `Unmapped` (was `Unmapped Unmapped`).

## Verification done

- `python -m unittest discover tests`: 8 tests pass (2026-10-20, with negative decimal fix and `Info` sheet checks).
- Scripted GUI test (2026-10-20, synthetic workbook): preview, Warnings view, sorting, unmapped filter, light/dark switch, user guide, export and Open buttons checked on screenshots.
- Full run on `data sheet classic FP SCA2.xlsx` with a test mapping: 600 tables, sheets `Summary_Per_Animal`, `Group_Statistics`, `T1`–`T21`, `Processing_Log`; 595 summary rows; 7 warnings (5 duplicates, unmapped tables, 1 unused mapping ID).
- GUI smoke test (scripted, real window): window 980×820, preview fills 600 rows, export succeeds, theme switch works.

## Open items / ideas

- Stride Width `(L)` and `(R)` are merged into one column (original design). Add a side column if the side matters for analysis.
- The decimal-fix rule misses damaged values that stay below the threshold. Possible upgrades: flag whole-number metric values (real averages almost never are), or check `cell.number_format` for a thousands format (`#,##0`) that Excel applies when it misreads the separator. Needs a real damaged file to confirm.
- The mapping file must have a header row; a file without one loses its first subject. Consider header detection.
- No GUI-level automated tests in the repo; the smoke test was run ad hoc.
- Optional: long-format export (one row per animal × timepoint × metric) for R/SPSS mixed models.
