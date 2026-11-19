# Running the invoice checks

Use the repository download, or clone `https://github.com/Ardholt/excel-vba-casebook`, then open `invoices-checked-reconciled/`. Docker is the only host prerequisite. `./run.sh` builds the shared image from [`../docker/Dockerfile`](../../docker/Dockerfile), tagged `excel-vba-casebook`, and invokes `scripts/inner.sh`. The first image build needs access to the image/package sources; processing runs with networking disabled. No host Python or Office installation is required.

The exact execution versions and times are recorded in [runtime.json](../evidence/runtime.json). The [Python version constraints](../requirements.txt) record the tested package versions. The shared Dockerfile defines dependencies; this case does not install a separate environment. To pin a reproduction, use the same repository revision and compare its versions with that record. A rebuild against newer package sources is not promised to be byte-identical.

## Replacement workbook

Put `input.xlsx` in the input directory. The output directory must be separate, outside the input directory and not its parent. Paths with spaces are accepted. Run:

```bash
./run.sh --input "your input folder" --output "your output folder"
```

Both the case source and input mount are read-only during processing. Only the requested output directory receives run files. Existing output files with the same names are replaced; unrelated files are kept. A repeat run replaces the reports instead of appending records.

Each sheet requires the following headers. Their order may change, extra columns are ignored, and the header row may occur within the first 25 rows. Entirely blank rows are ignored. Every nonblank record below the header is accounted for, including invalid records.

| Sheet | Required headers | Meaning |
|---|---|---|
| Lines | `invoice_id,line_id,quantity,unit_price,tax_rate` | Quantity may be fractional or negative. Price is nonnegative. Tax is a decimal fraction between zero and one. |
| Invoices | `invoice_id,declared_total` | One header per invoice; declared total has at most two decimal places. |
| Payments | `payment_id,invoice_id,amount` | Signed amount with at most two decimal places. Negative payments are refunds. |

Use value cells, not input formulas. Store IDs as text to retain leading zeroes. IDs are trimmed and case-sensitive; invoice IDs are unique, line IDs are unique within an invoice and payment IDs are unique across the batch. Identifiers are limited to 80 characters. All amounts are GBP in this case; there is no exchange-rate conversion. The schema and validation are implemented in [reconcile.py](../scripts/reconcile.py).

A tax rate of `0.20` means twenty percent. A negative quantity creates a credit line. Round each line’s quantity × unit price to cents, half-up; then round that net × tax rate to cents, half-up. Sum the rounded net and tax amounts. The declared difference is **declared minus calculated**, and the balance is **calculated gross minus allocated payments**. These rules apply equally to negative values.

Every occurrence of a repeated payment ID is held, whether amounts agree or conflict. A unique payment is allocated only when its invoice has a complete, unique header and valid lines. Unknown invoice references stay unallocated. No guessed allocation is made. Multiple unique payments with equal amounts may both be valid.

Invalid or repeated invoice/line data holds the entire affected invoice so a partial sum cannot masquerade as a complete total. Other invoices continue. Invalid payment amounts remain in the exception list; they are excluded from numeric totals and counted separately. Source row numbers refer to the original workbook, including the header offset. CSV exports prefix formula-like text identifiers with an apostrophe for safe opening in spreadsheet applications.

## What the files mean

| File or workbook tab | What it contains |
|---|---|
| `checked-ledger.xlsx` / Summary | Batch result and signed totals. |
| Payment list | Plain-language list of paid, part-paid, unpaid and credit amounts, with review notes. The opening rows show different payment outcomes; the full list retains every checked invoice. |
| Ledger | Calculated net, tax and gross; declared amount and difference; allocated payments and signed balance. |
| Line maths | Source line inputs and independently checked calculation formulas. |
| Payments | Every payment row, with allocated, duplicate_id, unallocated or invalid status. |
| Exceptions | Source sheet/row, record ID, reason and what requires review. |
| `invoice-views.xlsx` | A readable page per complete invoice, in input order. Each page identifies its invoice. |
| `invoice-views.html` | The same invoices with quantities, prices and rates; download and open locally to read or print. |
| `lines.csv`, `invoices.csv`, `payments.csv`, `exceptions.csv` | Row-level exports for audit and further processing. |
| `report.json`, `reconciliation.md` | Machine-readable and human-readable counts, amounts, exclusions and conservation checks. |
| `runtime.json` | Software versions, input hash, executed checks, timings and limitations for this run. |

These are generated reports. Edit the source workbook and rerun the command to incorporate changes; adding rows to an output workbook does not rerun payment matching. A negative balance is a credit or overpayment, a positive balance is due, and zero is settled. Declared-total differences do not alter the calculated amount or hide from the review list.

## Checks and rendering

The runnable code is [scripts/inner.sh](../scripts/inner.sh). It builds with Python/openpyxl, recalculates with LibreOffice, compares every formula result and runs [the tests](../tests/test_reconcile.py). [report.json](../output/report.json) separates identified business exceptions from failed processing checks. A schema or execution failure exits nonzero; it cannot print a successful result.

The page images are actual sheet renders. Single-sheet copies are used for export because the renderer can include every worksheet in a workbook. The sources are the original Overview, output Payment list (opening rows) and invoice view, with their sheet names and provenance in [media evidence](../evidence/media.json). The hero uses the shared casebook layout. No photographic assets are used; [asset notes](ASSETS.md) describe the typefaces.

For local page captures, the remote CI badge is rendered as a plain link to the test results; the README retains the required workflow badge.

The [page-capture script](../scripts/preview_page.py) produces phone and desktop README renders locally, not captures of the published GitHub or Pages site. The published page and Docker entry point have their separate repository acceptance checks.
