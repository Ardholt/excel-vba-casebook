# Sales dashboard out of date: £960.54 now counted

<p align="center">
  <img src="media/hero.png" width="600" alt="Updated sales summary: recent months, including January at £960.54, and every product category, including Repairs at £325.50.">
</p>

Latest months and every category shown; [see the full summary](images/breakdown.png).

**Paid sales rise from [£132,141.82 to £133,102.36](output/reconciliation.md) with the new orders included and refunds taken off. Refreshing again leaves the total unchanged.**

<sub>Case 04 of 8 in [Excel & VBA casebook](../README.md)</sub>

## What was wrong

- The sales table includes new orders, but the before dashboard shows the earlier total.
- Refunds needed taking off, while unpaid and cancelled orders needed keeping out of paid sales.
- Repeated order numbers and incomplete entries needed checking before they could be counted.

## What I did

- I rebuilt the refresh to include new orders in the sales total and month/category summaries.
- I deducted refunds and showed money still owed separately.
- I put repeated order numbers and incomplete entries on a review list, with a reason for each.
- I checked that refreshing again keeps the figures unchanged and leaves the original sales list intact.

## Proof

- **A full shop year:** [1,320 sales-list entries across 12 months](output/report.json), followed by [12 new entries](input/new-sales.xlsx).
- **New sales included:** [£132,141.82 → £133,102.36](output/reconciliation.md). The added January total is [£960.54](output/report.json); the new Repairs category contributes [£325.50](output/report.json).
- **Nothing lost or counted twice:** [1,332 entries = 1,300 accepted + 32 for review](output/reconciliation.md). [The review record](output/audit.csv) explains every decision.
- **Money still owed stays separate:** [£13,055.32](output/report.json). Cancelled orders do not increase paid sales, and refunds reduce them.
- **Refresh again, same result:** [the added orders change the totals once; another refresh leaves them unchanged](evidence/vba-run.json).

**Files:** [before the addition](output/vba-before.xlsx) · [finished dashboard](output/vba-after.xlsx) · [repeat refresh](output/vba-repeat.xlsx) · [workbook with refresh code](output/sales-dashboard.xlsm) · [original sales list](input/input.xlsx) · [totals and checks](output/reconciliation.md).

## Try it yourself

Optional technical check for a developer, from this folder with Docker running:

```bash
./run.sh
```

Creates the reference dashboard and row-by-row checks, then prints PASS. [Setup](DEPENDENCIES.md) · [replacement input format](input/SCHEMA.md) · [VBA source and setup](vba/README.md).

## How it is tested

[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)

Run in: [LibreOffice Calc 24.2.7.2 (headless)](evidence/vba-run.json), with the refresh macro checked against an independent reference calculation. **Microsoft Excel: not yet run.**

- [24 business tests](output/tests.json) check totals, refunds, unpaid/cancelled orders, duplicate IDs, invalid entries, moved headers and repeat runs.
- [The actual VBA checks](evidence/vba-run.json) compare every accepted row, audit decision, source value, summary and chart range before and after appending rows, then on repeat. Saved workbooks match too.
- The headless entry calls the same refresh core but does not exercise button installation (`OnAction`), the button click or the interactive error dialog. The supplied workbook tabs are required. [Formula recalculation checks](evidence/runtime.json) are recorded separately.

---

[£45 fixed, up to 5,000 rows, 3 working days.](https://ardholt.com/p/prices/)

More rows or pages: I quote the difference in writing before I start. Order through the platform.
