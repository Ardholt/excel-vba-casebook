# Run the cleanup on your workbook

Download this repository as a ZIP and extract it, or clone it. Install Docker with its engine running. Open a terminal in this case folder and run:

```bash
./run.sh
```

The command builds the [shared container](../docker/Dockerfile), writes the cleaned workbook and evidence to `output/`, and prints the checks followed by PASS or FAIL. It needs no host Python or spreadsheet application. The image build needs access to its package sources; the case itself runs with networking disabled. [Dependency instructions](DEPENDENCIES.md)

For your own files, use separate input and output folders; spaces are supported:

```bash
./run.sh --input "../customer exports" --output "../cleaned results"
```

Put `input.xlsx` and `rules.json` in your input folder. Your source files are mounted read-only. The output folder receives generated files; running again replaces those generated files. Keep your originals outside it. A run with your own inputs writes its evidence to that output folder; the page's pictures show the included inputs.

**Input format.** The workbook needs a sheet named `Records` with the headers `row_id`, `customer_id`, `name`, `date`, `amount`. Headers may move across columns or below title rows. Extra columns do not take part in matching and are not carried into Records; keep the original workbook. Completely blank rows are ignored. A partly filled row is assessed and, if invalid, retained in Rejected. [Parser and output writer](scripts/normalize_records.py)

- Store both IDs as text to preserve leading zeros. IDs are case-sensitive. Source sheet row numbers identify the audit entries even when row IDs repeat.
- `rules.json` must contain `{"date_order":"DMY"}` or `{"date_order":"MDY"}`. Dates accept year-month-day, slash dates with a full year, or workbook date cells without a time. No locale guessing.
- Amounts use a decimal point and optional comma thousands separators. Currency signs £, $ and € are accepted; brackets mean a negative amount. Inputs must use one currency: signs do not trigger conversion. Decimal-comma amounts, invalid grouping and excess decimal precision are rejected. [Parsing tests](tests/test_cleanup.py)
- Matching uses trimmed customer ID, whitespace-normalized and case-insensitive name, parsed date and amount. Similar spellings are not merged automatically. Blank required fields, numeric IDs and formula inputs are rejected with reasons. [Decision rules](scripts/normalize_records.py)


## Output tabs

Records is the usable table. Duplicates links removed rows to retained rows. Rejected preserves records needing review. Audit accounts for every nonblank source row; its value columns show normalized values for valid records and original values for rejected records. Changes lists each altered field. CSV files contain the same audit data; import their columns as text when inspecting literal source values.

## Rebuild the worksheet pictures

After the normal run, execute `python3 scripts/render_assets.py` inside the shared container, with the repository mounted so `../tools/hero/` is available. It rebuilds the proof workbook, before/after PNGs, two-frame GIF and hero from the current input and output. `python3 scripts/generate_input.py` restores the included input workbook before a full rebuild.
