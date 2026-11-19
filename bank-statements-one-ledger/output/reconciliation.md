# Statement reconciliation

810 input records; 810 ledger records; 0 rejected records.

| Account | Opening | Credits | Debits | Calculated close | Declared close | Difference | Status |
|---|---:|---:|---:|---:|---:|---:|---|
| ACCT-A | 1000.00 | 18366.36 | 8558.08 | 10808.28 | 10808.28 | 0.00 | OK |
| ACCT-B | 500.00 | 18218.50 | 8901.46 | 9817.04 | 9817.04 | 0.00 | OK |
| ACCT-C | 750.00 | 17959.64 | 8805.14 | 9904.50 | 9904.25 | -0.25 | REVIEW |

All money is GBP. Difference = reported minus calculated. Source order is retained.

## Exceptions

| Account | Source | Row | Check | Expected | Reported | Difference |
|---|---|---:|---|---:|---:|---:|
| ACCT-C | layout-c.csv | 271 | running_balance | 9904.50 | 9904.25 | -0.25 |
| ACCT-C | accounts.json |  | closing_balance | 9904.50 | 9904.25 | -0.25 |

Rejected records, if any, are in rejected.csv and report.json. An incomplete account cannot receive an OK status.
