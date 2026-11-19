# Files, inputs and the macros

The downloadable macro workbook contains the baseline input. The command-line route processes the input folder and produces a separate reference result. Neither route sends letters.

## Input schema

Put `input.xlsx` and `template.txt` in your input folder. The workbook must contain these sheets. Headers may move down or change column order; spelling must match. Fully blank rows are skipped. Extra columns are ignored in outputs.

| Sheet | Required columns | Rules |
|---|---|---|
| Data | record_id, group, value | Nonblank fields; numeric value. Every row is retained, including repeated IDs. |
| MergeA, MergeB | record_id, group, value | Same required headings; any order. MergeA rows precede MergeB rows. Repeated rows remain. |
| RenameMap | old, new | One request per nonblank row, applied in order. |
| Recipients | recipient_id, display_name, amount | Unique IDs ignoring case; IDs use letters, digits, underscore or hyphen, up to 64 characters; Windows reserved filenames are rejected. Names must be nonblank ASCII text; numeric amounts use at most two decimal places. |

Keep identifiers as text when leading zeroes matter. Supply values rather than formulas; paste calculated values before running. Sheet names beginning `__MP_` are reserved for generated results.

`template.txt` supports only `{{recipient_id}}`, `{{display_name}}` and `{{amount}}`. Template and recipient names use ASCII text for consistent plain-text output across the tested VBA file interface. Values are substituted once, so token-like text in a name stays literal. Letters end with a newline; amounts have a period and two decimal places. This workbook uses GBP in its template; edit that wording for your own currency.

Rename rules: source names are matched without regard to case. Missing sources, repeated source requests, names exceeding Excel's name limit, forbidden characters, leading/trailing apostrophes, the reserved History name, protected input sheets and existing destination names are rejected. No other sheet is overwritten. Requests are sequential, so a swap is rejected as a collision. The rename log reports each decision. A second run leaves the new tabs in place and reports the old source names as missing.

## Outputs

| File or tab | Meaning |
|---|---|
| after.xlsx | Reference summary with original source tabs, a split file list and the other action results. |
| __MP_Split | Group-to-file lookup and row count. |
| split/group-001.xlsx, … and split-workbooks.zip | One workbook per group, in first-seen group order; one Data sheet per file. The reference also bundles the files in a ZIP. Numbered filenames avoid unsafe or clashing group labels. |
| __MP_Merged | All eligible rows, aligned as record_id, group, value. |
| __MP_Rename | Original name, requested name and decision. |
| __MP_Letters | Recipient, formatted amount and exact letter text. |
| Results | Reference run totals. |
| split.csv, merged.csv, renames.csv | Readable output and audit tables. |
| letters/ and letters.zip | One text file per recipient ID, also bundled for download. |
| report.json, run.txt | Counts, exclusions, checks and reference runtime. |

Use a dedicated output folder. The reference replaces its generated letter files and numbered `split/group-*.xlsx` workbooks on each run; do not store unrelated text files in `output/letters`. The VBA replaces matching recipient files and removes files recorded in its previous letter log, inside its own letters folder only. Keep unrelated files elsewhere. The VBA split replaces numbered files for the current run and removes obsolete files recorded in its previous split log; unrelated filenames are preserved. Both routes leave input files unchanged; VBA changes the open macro workbook only when you invoke its actions.

## Run the workbook manually

Review the readable [split](vba/SplitRows.bas), [merge](vba/MergeTables.bas), [rename](vba/RenameSheets.bas) and [letter](vba/MakeLetters.bas) modules before enabling macros. Open [macro-pack.xlsm](macro-pack.xlsm) in a writable folder, enable macros for this workbook and use the Macros dialog to run `SplitRows`, `MergeTables`, `RenameSheets`, and `MakeLetters`. They are independent actions; there is no auto-open procedure. Split writes one workbook per group in `output/split` beside the macro workbook and a file list in `__MP_Split`. Merge writes its result tab, and rename changes that workbook’s tabs. MakeLetters writes `output/letters` beside the workbook. Save a separate copy after reviewing the results.

To use your own data in the macro route, replace the source tabs in a fresh copy of the macro workbook, retaining their required names. Copy your template text into `Template!A1`. Re-run the chosen action. New groups create additional numbered workbooks. The summary tabs use a reserved prefix and are replaced on repeat runs. Do not put your own data in those tabs.

If the packaged workbook cannot be opened cleanly in Excel, import all four `.bas` files into a copy of the input workbook in the Visual Basic Editor, add the Template sheet, set A1 to your template text and save as `.xlsm`. The embedded package combines the four source bodies in Module1; import the individual files only into a clean workbook, to avoid duplicate procedure names.

See [How it is tested](README.md#how-it-is-tested) for the execution status and the boundary of the validated path. 

## Container and reproduction

The shared `../docker/Dockerfile` is the dependency definition used by `run.sh`; it is based on Ubuntu 24.04 with LibreOffice, Python, openpyxl, XlsxWriter and pytest. The exact executed versions are in [runtime.json](evidence/runtime.json). Docker itself was not run on the development host; the shared CI owner checks the wrapper.

`bash scripts/inner.sh` is the same reference/check/test sequence inside the shared image. `python3 -m pytest -q tests` runs the business tests. To rebuild and test the VBA candidate in the image with its isolated packaging dependency:

```bash
/opt/vba/bin/python scripts/package_vba.py
python3 scripts/build_candidate.py
python3 scripts/run_vba.py
```

The VBA runner enables executable import only in a disposable LibreOffice profile for this reviewed workbook. It compares the imported source and each action's outputs, and exports snapshot workbooks through LibreOffice's file interface. For split workbooks the macro itself uses a LibreOffice branch with the native OOXML export filter. The runner opens those files, checks their ZIP format and exact rows, and tests new groups, repeated rows, stale-file cleanup and empty input. The Excel branch uses Workbook.SaveAs with FileFormat 51.
