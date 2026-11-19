# Paid invoice, wrong total

The payment covers the lines below, but the invoice asks for **£2.35 too much**. The check keeps that difference open for review.

<img src="media/hero.png" width="560" alt="INV-0017 source beside Ledger row 18. Net £227.49 + tax £10.63 = gross £238.12. Stated £240.47 − calculated £238.12 = £2.35 to review. Gross £238.12 − recorded payment PAY-0017 £238.12 = £0.00 remaining. Payment covered; invoice total still needs checking.">

Trace this check: [source invoice and payment](input/input.xlsx) · [ledger row](output/invoices.csv) · [line calculations](output/lines.csv) · [£2.35 review entry](output/exceptions.csv).

**[120 invoices checked; 7 totals need review.](output/reconciliation.md)** The checked workbook keeps the total check separate from the payment check.

<sub>Case 06 of 8 in [Excel & VBA casebook](../README.md)</sub>

## What was wrong

- Some invoice totals did not match the amounts on their lines.
- Repeated payment references could make an invoice look more fully paid than it was.
- Part-payments, money returned and extra payments made it hard to see what was still owed.

## What I did

- I added up each invoice’s lines and tax, then checked its stated total.
- I matched payments to invoices and set repeated entries aside for checking.
- I made a clear list of what is paid, what is still owed and where there is money to return or use as credit.
- The `.xlsx` workbooks contain that payment list and a separate list of entries to review.

## Proof

The [Payment list in the checked workbook](output/checked-ledger.xlsx) shows each invoice’s payment position. Entries with questionable payments are marked **Check payment**, so they can be reviewed before asking someone to pay again.

| What you can see | Result from these files |
|---|---|
| Invoices checked | [120 invoices, with 360 lines](output/report.json) |
| Totals needing review | [7, overstated by £16.45 altogether](output/reconciliation.md) |
| Payments matched to invoices | [£40,326.02 across 119 entries](output/reconciliation.md) |
| Repeated payment references | [6 entries sharing 3 references, set aside for checking](output/report.json) |
| Payments with no matching invoice | [4 entries, £183.15 altogether](output/reconciliation.md) |

Every [payment entry is kept in the payment file](output/payments.csv). The repeated entries need checking against the payment records before anyone can say how much of them was actually received. They are not counted as additional money received.

### Follow the invoice in the first image

The image is a compact view of **INV-0017** from the supplied workbook beside its checked ledger entry. It shows the full check, without cropping a longer payment list.

In [input.xlsx](input/input.xlsx), the invoice is on **Invoices row 18**, its lines are on **Lines rows 50–52**, and **PAY-0017** is on **Payments row 18**. The corresponding [checked workbook](output/checked-ledger.xlsx) entry is **Ledger row 18**; its `total_check` is `REVIEW` even though its payment balance is `SETTLED`.

| Line | Net calculation (£) | Tax calculation (£) | Gross (£) |
|---|---|---|---:|
| L01 | 4 × 21.93 = 87.72 | 87.72 × 0% = 0.00 | 87.72 |
| L02 | 5 × 23.10 = 115.50 | 115.50 × 5% → 5.78 | 121.28 |
| L03 | 1 × 24.27 = 24.27 | 24.27 × 20% → 4.85 | 29.12 |
| Total | 227.49 | 10.63 | 238.12 |

These values are in [lines.csv](output/lines.csv). Net and tax are each rounded half-up to pennies per line before summing. The [invoice ledger](output/invoices.csv) records **£227.49 + £10.63 = £238.12**, and **£240.47 − £238.12 = £2.35** to review. The [payment file](output/payments.csv) records **PAY-0017: £238.12**, giving **£238.12 − £238.12 = £0.00** remaining against the calculated total. The [exception](output/exceptions.csv) remains open; the stated invoice has not been silently changed.

**Download:** [original workbook](input/input.xlsx) · [checked workbook with payment list](output/checked-ledger.xlsx) · [invoice views](output/invoice-views.xlsx) · [entries to review](output/exceptions.csv) · [full reconciliation](output/reconciliation.md)

## Try it yourself — optional developer check

A developer can reproduce the checks with Docker from this case folder:

```bash
./run.sh
```

It rebuilds the files and prints `PASS` or `FAIL`. [Setup and instructions for another input workbook](docs/USAGE.md).

## How it is tested

The [verification record](evidence/verification.json) is a historical run; its source hashes identify that snapshot, not the current rendering/page sources.

[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)

The `.xlsx` workbooks were produced with [Python 3.12.3 and openpyxl 3.1.2](evidence/runtime.json), then checked with [LibreOffice 24.2.7.2: all 1,920 recalculated formula cells matched](evidence/runtime.json).
[31 tests passed](evidence/runtime.json), covering rounding, refunds, part-payments, repeated payment references and payments with no matching invoice.
The tests also check that no records disappear, the input stays unchanged and repeat runs give the same result. [Tests](tests/test_reconcile.py) · [replacement-input checks](evidence/replacement-input.json) · [technical rules and workbook tabs](docs/USAGE.md).

---

The files above show [120 invoices](output/report.json). The offer below covers a smaller batch; for your larger pile, I quote the total through the platform before I start.

[£10 fixed, up to 20 invoices, net/VAT/gross that checks itself, 2 working days.](https://ardholt.com/p/prices/) Order through the platform.

This case at its size (120 invoices): £60 fixed, £10 for each 20 invoices.

More rows or pages: I quote the difference in writing before I start.
