# Sales input

Put `input.xlsx` in the directory passed to `--input`. It must contain a `Sales` sheet. Required column names are `sale_id,date,category,amount,status`; column order may change and the header may appear within the first fifty rows. Each required name must occur once. Extra columns are ignored.

- `sale_id`: nonblank identifier; surrounding spaces are removed. IDs are case-sensitive. Every occurrence of a duplicated ID is rejected for review, including conflicting duplicates. No automatic winner is selected.
- `date`: literal ISO text, `YYYY-MM-DD`, a valid calendar date from 1900 onward. Actual spreadsheet date cells are not the input format; store the ISO text.
- `category`: nonblank text. Case and punctuation are significant after trimming. Categories containing wildcard characters are treated literally.
- `amount`: signed decimal with at most two decimal places; numeric cells or decimal text, a period for text decimals, no currency symbols or grouping separators. Absolute value must not exceed 999999999. No rounding of invalid precision. This workbook uses GBP throughout; do not mix currencies.
- `status`: exactly `Paid`, `Pending` or `Cancelled` after trimming. Negative paid amounts are returns and remain part of paid transaction count.

Completely blank required fields are ignored. A partially completed row is reviewed. Formulas and error cells in required fields are rejected; export literal values first. Pending totals are shown separately. Cancelled rows remain in the audit and accepted-data sheet but do not contribute to paid revenue. Review priority is missing ID, duplicate ID, formula/error, date, category, amount, then status; one reason is recorded per rejected row.

`before.xlsx` is an optional earlier snapshot used only to produce the before dashboard and comparison. `new-sales.xlsx` holds the additional sales rows for the desktop run; it is not loaded by the command. `input.xlsx` already contains those appended rows. The included inputs can be rebuilt with `python3 scripts/generate_input.py` inside the supplied container.

The audit CSV prefixes formula-like text with an apostrophe so spreadsheet applications open it as text. Numeric amounts remain numeric text. The workbook audit stores those fields as literal text.

The source files are never modified. Choose a separate output directory. The run overwrites its named reports and workbooks there; unrelated files are left alone. The run is a Python reference, not a macro runner.
