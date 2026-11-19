# Execution evidence

[The report](../output/report.json) and [reconciliation](../output/reconciliation.md) account for the full year and appended sales. [The audit](../output/audit.csv) preserves every source row and exclusion reason.

[Runtime](runtime.json) records the Python reference, software versions and actual LibreOffice formula recalculation. [Business tests](../output/tests.json) record the checks run again after the final page render.

[The actual VBA execution](vba-run.json) ran the packaged Module1 in LibreOffice 24.2.7.2 headless. The baseline, appended rows and repeat refresh each matched the reference for every accepted row, audit decision, source value, paid/pending total, month/category total and chart range. The saved outputs were checked independently as well. [Packaging](macro-package.json) records source identity and the separate opening check.

The headless entry calls the same refresh core as the button, while omitting shape OnAction assignment, the click and interactive failure dialog. The supplied fixed tabs are required; recreation of deleted tabs is not supported by this workflow. See the [execution status](../README.md#how-it-is-tested).


[Sheet captures](images.json) identify the actual VBA output workbooks. [Hero evidence](hero.json) records the visible month/category rows and their saved-workbook source. [Page evidence](page-runtime.json) records browser versions and capture hashes. The [phone first screen](page-preview/390-first.png) shows the title, recent months and every category. The [offline rendering note](page-source.json) explains why the remote badge image is omitted from the captures. These are local previews, not visits to the published site.

[Validation](validation.json) records the successful final executions and call totals. [Final file checks](file-checks.json) cover identifying properties, metadata, vocabulary and local links.
