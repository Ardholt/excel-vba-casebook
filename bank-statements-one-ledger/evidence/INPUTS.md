# Run the statement check

Download and extract the whole repository using GitHub's Code menu, or clone it. Keep this case beside the shared `docker` folder. Install Docker with its daemon running, and open a Bash terminal in this case folder. The command needs no host Python, Office installation, keys or private files.

```bash
./run.sh
```

The first image build needs package-repository access; the case container runs without network access. The runner builds `excel-vba-casebook` from `../docker`, processes the statements, checks the saved workbook with LibreOffice, runs the business tests, and prints PASS or FAIL. [Runtime and dependency details](REPRODUCING.md).

## Replacement statements

Use the documented filenames and put them with `accounts.json` in a separate directory:

```bash
./run.sh --input "my statements" --output "checked ledger"
```

Input and output directories must be separate and non-nested. Sources are read-only; only the chosen output directory is written. A repeat run replaces named outputs rather than appending transactions. [Executed command-line checks](input-options.json) cover changed input values, paths with spaces, unchanged source bytes and matching repeat results.

## Input format

Headers may move down the sheet/file or appear in another column order. Extra columns are ignored. Blank rows are skipped. A nonblank invalid row is listed in the rejected report. CSV is UTF-8. Dates accept ISO `YYYY-MM-DD` or slash dates using the declared `DMY`/`MDY` order. [Reader and validation rules](../scripts/ledger.py).

| File | Required sheet / columns |
|---|---|
| `layout-a.csv` | `id,date,description,debit,credit,balance` |
| `layout-b.xlsx` | `Transactions`: `id,booking_date,details,amount,balance` |
| `layout-c.csv` | `reference,date,memo,amount,direction,balance` |

Money is a plain decimal number with at most two decimal places, without currency signs or thousands separators. Debit/credit blanks in layout A mean zero; both cannot be nonzero. Layout B uses signed amounts. Layout C uses unsigned amounts with `IN` or `OUT`. Workbook inputs contain values, not formulas. [Parsing tests](../tests/test_ledger.py).

`accounts.json` maps each filename to `account`, `currency`, `opening_balance`, `closing_balance`, and `date_order`. Use distinct anonymous account codes, one statement per account. All money is GBP. The settings supply opening and closing balances; these are not guessed from transactions. [Included configuration](../input/accounts.json).

Equal-looking payments are preserved. A repeated identifier within an account is rejected for review; similar transactions on different accounts remain. When a row is rejected, later balances on that account are unverified and its close is marked incomplete; the command returns a failure code. Valid records with a balance mismatch return a successful diagnostic and REVIEW status. [Tests](../tests/test_ledger.py).

## Output tabs

- **Balance checks:** per-account result.
- **Ledger:** signed transactions, source references and every running balance check.
- **Reconciliation:** date coverage, openings, money in/out, calculated and declared closes.
- **Exceptions:** failed checks with expected and reported values.
- **Rejected:** invalid records and reasons.
- **Trace a balance:** the first discrepancy, or final transaction if none exists.
- **Closing balance proof:** a compact reported/calculated comparison.

The input overview contains all original source columns and row numbers, a **Statement layouts** view and a **Source balance proof** view. The compact views supply the page's actual sheet captures. [Source workbooks and outputs](../README.md#proof).
