# Sales reconciliation

Paid revenue: £133102.36 across 1149 paid transactions.
Pending: £13055.32 across 106 transactions; excluded from paid revenue.
Cancelled: 45 rows (£6069.30) excluded from paid revenue.
Conservation: 1332 input rows = 1300 accepted + 32 rejected.

Duplicate policy: exclude every occurrence of a repeated nonblank ID. Review and correct the source; no first-row guess.

| Reason | Rows |
|---|---:|
| Duplicate sale ID: all occurrences excluded | 24 |
| Invalid ISO date | 3 |
| Invalid amount: decimal with up to two places required | 2 |
| Invalid status | 1 |
| Missing category | 1 |
| Missing sale ID | 1 |

| Paid month | GBP |
|---|---:|
| 2025-01 | 6973.95 |
| 2025-02 | 7541.38 |
| 2025-03 | 8803.11 |
| 2025-04 | 9868.83 |
| 2025-05 | 10521.09 |
| 2025-06 | 10121.77 |
| 2025-07 | 11545.58 |
| 2025-08 | 12651.48 |
| 2025-09 | 10673.51 |
| 2025-10 | 11710.97 |
| 2025-11 | 14113.52 |
| 2025-12 | 17616.63 |
| 2026-01 | 960.54 |

| Paid category | GBP |
|---|---:|
| Garden | 23358.10 |
| Hardware | 21187.72 |
| Household | 21492.51 |
| Lighting | 19773.25 |
| Paint | 23533.34 |
| Repairs | 325.50 |
| Tools | 23431.94 |

Before: £132141.82; after: £133102.36; change: £960.54.

[Machine-readable figures](report.json) · [Row-by-row audit](audit.csv) · [Runtime](runtime.json)

All totals are derived from accepted rows using decimal arithmetic, then compared with the recalculated workbook.
Reference execution does not establish desktop macro execution.
