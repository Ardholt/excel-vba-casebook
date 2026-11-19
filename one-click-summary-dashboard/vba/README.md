# Refreshing the sales dashboard

The [downloadable workbook](../output/sales-dashboard.xlsm) includes the **Refresh sales** Forms button and its VBA. It opens with the completed sales list and the £133,102.36 result. The [recorded LibreOffice run](../output/vba-run.json) pressed the installed button before and after the additions, pressed it again, saved the macro workbook, then reopened it and pressed its button again.

## For staff

1. Open `sales-dashboard.xlsm` in desktop Excel. Review [SalesDashboard.bas](SalesDashboard.bas) and enable macros for this workbook.
2. Append new entries to **Sales**, keeping the existing columns and one header row. Use [the input format](../input/SCHEMA.md).
3. On **Dashboard**, click **Refresh sales**, below the total. The total, both breakdowns and both charts rebuild from the sales list.
4. Check for **“Checked: all source rows accounted for”** below the total. Read **Audit** for rejected entries and their reasons; correct **Sales**, then refresh again. Save when finished.

The refresh leaves **Sales** unchanged. **ValidSales**, **Audit** and **Dashboard** are dedicated output sheets whose contents are rebuilt. Keep these tabs in place and put personal notes elsewhere. All occurrences of a duplicate order number stay excluded until corrected. Refunds reduce paid sales; pending and cancelled entries stay out of the paid total.

A failed refresh writes **“REFRESH FAILED - correct source and retry”** and displays the error. Resolve the cause and refresh successfully before using the figures.

## Rehearse the sample addition

The finished download already includes the new rows. To reproduce the before/after sequence in your own copy:

1. Open a copy of the button workbook. In **Sales**, replace the existing data rows with the rows from [before.xlsx](../input/before.xlsx), keeping one header. Click **Refresh sales**: £132,141.82.
2. Append the 12 data rows from [new-sales.xlsx](../input/new-sales.xlsx), without its header. Click **Refresh sales**: £133,102.36, including January at £960.54 and Repairs at £325.50.
3. Click again: £133,102.36. Compare [the reconciliation](../output/reconciliation.md) and save.

## Rebuild or install manually

`../run.sh` packages the current module and a Forms button bound to `RefreshDashboard`, then runs the button checks and saves fresh outputs. No separate installation macro is needed.

For manual assembly in Excel, open [before-dashboard.xlsx](../output/before-dashboard.xlsx), save it as **Excel Macro-Enabled Workbook (.xlsm)**, and import [SalesDashboard.bas](SalesDashboard.bas) in the Visual Basic Editor using **File → Import File**. On **Dashboard**, insert a **Button (Form Control)** from the Developer tab, assign **RefreshDashboard**, label it **Refresh sales**, then save.

## Implementation and checks

The button invokes `Module1.RefreshDashboard`, which calls `RefreshSalesCore`. The module uses worksheet ranges, integer-cent arithmetic and charts. It has no automatic startup event, network call, file operation, ActiveX control, Dictionary or external scripting dependency. It validates rows before rebuilding, checks row conservation and paid/pending/month/category totals, and restores calculation, events and screen updating on success or failure.

[How it is tested](../README.md#how-it-is-tested) describes the actual run: LibreOffice imports the packaged Forms button and its `actionPerformed` binding, then the runner activates its `XAccessibleAction`. No direct call to a separate headless macro is used. Python independently verifies every accepted and audit row, source preservation, totals, chart ranges and saved files. [Package verification](../output/macro-package.json) compares extracted VBA with the source after saving.

The workbook project template comes from pyOpenVBA under its [MIT licence](PACKAGING-LICENSE.txt). The sales data, dashboard, button and module are authored for this case.
