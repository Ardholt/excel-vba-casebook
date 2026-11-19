# Refresh the folder with Power Query

This Power Query version has not been run yet: the recorded run used the VBA macro in LibreOffice Calc, checked against a Python reference (see [How it is tested](../README.md#how-it-is-tested)). These setup steps require an Excel installation with the folder connector and a blank-query editor; availability varies by edition and platform.

1. Keep the downloaded source workbooks together in `input/files`. Create a new Excel workbook outside that folder. Save it as `folder-query.xlsx`.
2. On a sheet named `Config`, put the full source-folder path in a cell. Select the cell and name it `SourceFolder` using the Name Box. Include spaces normally; do not add quotation marks around the path.
3. Open Power Query's blank-query editor. Name the query `FolderResult`. Open Advanced Editor, replace its contents with [CombineFolder.m](CombineFolder.m), and finish. Keep this record query as connection only.
4. Create a blank query named `Combined`. Its Advanced Editor contents are `let Result = FolderResult in Result[Combined]`. Load this table to a new sheet.
5. Create queries named `Manifest` and `Rejected` in the same way, replacing `[Combined]` with `[Manifest]` and `[Rejected]`. Load both tables so that skipped files stay visible.
6. Use Refresh All. Check the loaded tables against [the reference report](../output/report.json) and [the per-file manifest](../output/manifest.csv). A file rejected for missing columns must be repaired at source before it can enter the table.
7. Add a valid workbook to the folder and refresh again, then refresh once more without changing files. The query is written so that the new file enters once and a second refresh changes nothing; check both on your copy. Save the workbook after reviewing `Rejected`.

As written, the query scans the immediate folder, ignores temporary `~$` files and named outputs, selects the `Data` worksheet, trims headers and aligns by name. Optional columns are retained; absent values remain null. It keeps the original source row before removing entirely blank rows. Equal-looking records remain separate. It does not deduplicate, convert currencies or infer dates. Use the [input contract](../USAGE.md) for replacement exports.

