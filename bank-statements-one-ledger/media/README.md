# Where the opening proof comes from

The hero is a compact summary drawn from the case files, with selected columns rather than a screenshot of Excel. All money is in pounds. A, B and C in the account checks mean ACCT-A, ACCT-B and ACCT-C.

| Source row (header included) | Source money columns | Matching ledger row |
|---|---|---|
| [layout-a.csv](../input/layout-a.csv), row 3: A002 | debit 125.50; credit blank | ACCT-A, A002: −125.50 |
| [layout-b.xlsx](../input/layout-b.xlsx), Transactions row 3: B002 | amount −500.00 | ACCT-B, B002: −500.00 |
| [layout-c.csv](../input/layout-c.csv), row 3: C002 | amount 200.00; direction OUT | ACCT-C, C002: −200.00 |

The blank credit is shown as “—”. The matching rows are in [ledger.csv](../output/ledger.csv), with the source filename and row number retained. These three examples show the formats being combined; the [report](../output/report.json) records all 810 rows retained.

Every closing balance and status comes from [reconciliation.csv](../output/reconciliation.csv). A and B match; C reports £9,904.25 against £9,904.50 calculated, a −£0.25 difference requiring review. No amount has been adjusted to balance it.

Run `python3 scripts/page_pass.py` inside the case container to rebuild the hero and page previews. It reads the source and output files directly, checks the sample amounts agree, and records their hashes and displayed rows in [hero.json](../evidence/hero.json). Typography uses the container’s DejaVu Sans fonts.

The earlier worksheet captures remain available as [before.png](before.png) and [after.png](after.png); the complete [source layouts](../images/statement-layouts.png) and [account checks](../images/balance-checks.png) provide the wider worksheet views.
