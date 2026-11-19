# Reconciliation

Checked invoices: 120; line rows: 360; payment rows: 129.

Declared totals requiring review: 7. Credit notes: 10.

| Invoice totals (GBP) | Amount |
|---|---:|
| net | 40951.88 |
| tax | 3463.95 |
| gross | 44415.83 |
| declared_total | 44432.28 |
| difference | 16.45 |
| allocated | 40326.02 |
| balance | 4089.81 |

| Payment disposition | Rows | GBP |
|---|---:|---:|
| allocated | 119 | 40326.02 |
| duplicate_id | 6 | 2161.59 |
| unallocated | 4 | 183.15 |
| invalid | 0 | 0.00 |

All parsed payment rows: GBP 42670.76. Repeated IDs are not confirmed payments: their rows stay held, including the first occurrence.

A negative balance means credit owed or an overpayment. Balances are calculated against line totals, while declared differences stay open for review.

- all_line_rows_accounted_for: PASS
- all_invoice_rows_accounted_for: PASS
- all_payment_rows_accounted_for: PASS
- payment_money_conserved: PASS
- net_plus_tax_equals_gross: PASS
- gross_less_allocated_equals_balance: PASS

LibreOffice formula values compared: 1920; PASS. Input bytes unchanged: PASS.
