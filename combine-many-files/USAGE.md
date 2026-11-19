# Run and refresh the folder

Download the repository ZIP and extract it, or clone the repository. Open a terminal in `combine-many-files`. Install Docker with its engine running; the container supplies Python and LibreOffice, so the command needs no local Office installation. The shared build definition is [docker/Dockerfile](../docker/Dockerfile); measured versions are in [runtime.json](evidence/runtime.json).

```bash
./run.sh
```

The command creates a new empty `runs/run.XXXXXXXX` folder and prints its path on completion. It mounts the case and original inputs read-only, with only that new folder writable. It copies code and tests into the run folder, builds the reference table from scratch, packages and runs the VBA module in LibreOffice headless, compares the results, runs the tests and prints `PASS: combine checks`. Published downloads in `output/` remain available. Each repeat invocation gets a separate fresh folder. A rejected input file is a visible diagnostic; it does not make the valid exports disappear. Failed checks return a nonzero status.

To consolidate your own exports, keep input and output in separate, non-nested folders:

```bash
./run.sh --input "/path/to/branch exports" --output "/path/to/results"
```

This route processes your files with the Python reference and checks every resulting cell against those sources. It does not run a macro on your files. Input and case code are mounted read-only. The output folder must be empty or new; a nonempty folder is refused before Docker starts. Omit `--output` to create a fresh folder under `runs/`. It accepts paths containing spaces. Workbook values, dates and amounts must follow the contract below.

## Input contract

Use `.xlsx` workbooks in the immediate input folder. Each has a sheet named `Data`. The first row contains the exact required headers below, in any order; surrounding header spaces are removed. Each later nonblank row is a source record. Entirely blank rows are skipped without renumbering provenance.

| Header | Value |
|---|---|
| `record_id` | Identifier stored as text; leading zeros and repeated identifiers are preserved. |
| `date` | ISO date text, `YYYY-MM-DD`. |
| `category` | Category text. |
| `amount` | Numeric amount in a single consistent currency, with at most two decimal places. Credits may be negative. |

Extra columns are retained. Missing extra values stay blank. Headers must be unique and nonblank; `source_file` and `source_row` are reserved. A missing `Data` sheet, missing required column or unreadable workbook goes into the rejection report. A header-only sheet is an empty accepted structure and adds no rows. This is a value-consolidation task, not a date, amount or formula-cleaning service: use stored values, not formula cells or stale formula caches. No columns or equal-looking source records are discarded to make totals fit.

The temporary `~$` files, non-XLSX files and known outputs (`consolidated.xlsx`, `vba-result.xlsx`, `combine.xlsm`) are excluded. Keep other output filenames outside the source folder. Subfolders are not traversed.

## Output files and sheets

The paths below link to the published example. A default command run writes its own `output/`, `evidence/` and `combine.xlsm` inside the printed run folder. A custom-input run writes the reference files directly into its chosen result folder.

| File or sheet | What it shows |
|---|---|
| `output/consolidated.xlsx` → Combined | Required fields, filename, original row number and every optional field. |
| Manifest | One decision per file: accepted, empty, rejected or ignored; row counts and reasons. |
| Rejected | Files that could not enter the table. |
| Refresh check | A compact view of the measured totals. |
| `output/combined.csv` | The same rows in an open table format. |
| `output/report.json` | Counts, amount total, runtime, source hashes and conservation checks. |
| `output/vba-result.xlsx` | Cells produced by the actual VBA run in Calc, exported through UNO. |

A refresh rebuilds the whole table; it does not append yesterday's output. Reconcile money with the same currency and decimal convention across sources. The workbook does not silently convert units.

## Run the macro in Excel

See [How it is tested](README.md#how-it-is-tested) for the recorded execution and versions.

Open [combine.xlsm](combine.xlsm), review [the readable module](vba/CombineFiles.bas), and enable macros only for this trusted workbook. Enter the full local source-folder path in `Config!B2`. A relative folder starts beside the workbook. Use Excel's Macros menu to run `RefreshFolder`. Review `Config` status and the `Combined` and `Manifest` sheets. Run it again to refresh. Save a separate copy after checking the results.

If the embedded project cannot open without repair, create a workbook with `Config`, `Combined` and `Manifest` sheets, import `vba/CombineFiles.bas` in the VBA editor, set `Config!B2`, and save as XLSM. There is no auto-run or button; the macro performs no file deletion or network access. Sources are opened read-only with link updates disabled. Source macros are disabled while reading and the prior application settings are restored afterwards.
