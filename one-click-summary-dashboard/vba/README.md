# Desktop refresh setup

The actual packaged VBA refresh was run in [LibreOffice 24.2.7.2, headless](../evidence/vba-run.json). Baseline, appended sales and repeat refresh match the Python reference, including every accepted/audit row, both chart ranges, source preservation and saved workbook totals. `RefreshDashboardHeadless` calls the same `RefreshSalesCore` as `RefreshDashboard`, omitting button installation and the interactive dialog. No Dictionary or external scripting runtime is used.

See the [execution status](../README.md#how-it-is-tested). `OnAction` assignment and the button click were not exercised in the headless run.


Use Microsoft Excel for Mac or Windows with macros enabled for this workbook after reviewing the source. The module uses native worksheets, charts and a shape button; it has no automatic startup event, network call, file operation, ActiveX control or external scripting dependency.

Open [sales-dashboard.xlsm](../output/sales-dashboard.xlsm), which contains the before-state sales and the readable code in `Module1`. [Packaging and opening checks](../evidence/macro-package.json) verify the module source and baseline values with macros disabled; real Excel acceptance is still required. Run `InstallRefreshButton` once from the Macros dialog, then save.

If Excel rejects the candidate or requests repair, use this fallback: open [the before dashboard](../output/before-dashboard.xlsx), save a copy as **Excel Macro-Enabled Workbook (.xlsm)** named `sales-dashboard.xlsm`, open the Visual Basic Editor, and import [SalesDashboard.bas](SalesDashboard.bas) using **File → Import File**. Run `InstallRefreshButton` once from the Macros dialog. The imported module is named `Module1`. It adds the **Refresh sales** shape button. Save the workbook.

Click **Refresh sales** to rebuild from its `Sales` sheet. Append the data rows from [new-sales.xlsx](../input/new-sales.xlsx) to the bottom of `Sales` without repeating its header, then click again. Compare the dashboard with [the reconciliation](../output/reconciliation.md), and click once more to confirm the same result. Never append these rows to `input.xlsx`, which already includes them.

The source `Sales` sheet is untouched. `ValidSales`, `Audit` and `Dashboard` are required, dedicated derived sheets whose contents are rebuilt on refresh; keep these tabs in place and keep your own notes elsewhere. Review rejected rows in `Audit`, correct `Sales`, then refresh. All instances of duplicate IDs stay excluded until the IDs are resolved.

The macro parses and totals integer cents, then checks row conservation, paid/pending totals, and month/category totals before showing its successful check line. A failure shows a dialog and must be resolved before using the results. The supported source format and validation rules are in [the schema](../input/SCHEMA.md).

The candidate uses the embedded workbook-project template from pyOpenVBA under its [MIT licence](PACKAGING-LICENSE.txt). The sales data, dashboard and module are authored for this case.
