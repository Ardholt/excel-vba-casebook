# Customer register and review rules

The entry point is `./run.sh`. It needs Docker and the whole repository, including `../docker/Dockerfile`; no host Python installation or spreadsheet application is needed. The container's installed versions are recorded by each run in `output/runtime.json`. The shared container definition is maintained at repository level. [requirements.txt](requirements.txt) pins the Python dependencies used by the recorded container run. The tested versions are linked in [execution evidence](evidence/runtime.json).

## Running your own files

Download and unzip the complete repository, install Docker, then open a terminal in this case folder and run `./run.sh`. The command builds the shared container and checks the bundled register.

For another register, use `./run.sh --input "my register" --output "my results"`. Start without `confirmed.csv` to keep every customer separate. Review the suggestions, fill `confirmed-template.csv` with only the choices you approve, save it as `confirmed.csv` in the input folder and rerun against the original register. Never copy every suggestion into approvals automatically.

## Inputs

`input.xlsx`, sheet **Customers**, needs these unique header names, in any column order. A title or blank rows may appear above the headers. Fully blank data rows are ignored. Extra columns are ignored; this case carries only the fields below into the retained register.

| Header | Rule |
|---|---|
| customer_id | Nonblank unique text. Leading zeros are preserved. Numeric IDs and formula cells are rejected. |
| name | Nonblank text. Kept as entered except outer whitespace. Comparisons ignore case and repeated whitespace. |
| postcode | Text area code; blank is allowed but is not compared. The bundled register uses abstract area codes. |

`confirmed.csv` is optional, UTF-8, with `source_id,target_id` headers in either order. `source_id` is the old customer number; `target_id` is the customer number to keep. Both must exist in the original register. The source is linked to the target; the target's name and area code are kept. Confirmations need not appear among proposals: the owner's decision can supply a match the suggestion rules missed. An absent or header-only file applies no merges.

Each run reads original input only and replaces its named output files. Keep input and output folders separate, with neither nested inside the other. The supplied Docker command mounts input read-only. Paths containing spaces are supported; commas in directory paths are not supported by the Docker mount syntax used here.

## Proposal rules

Read [the matching source](scripts/customers.py) for the exact rule: normalize Unicode, case and whitespace; compare only records with the same nonblank normalized area code; propose a pair when the name similarity is at least **0.82**. Similarity is the character-sequence ratio from Python's `SequenceMatcher`, with automatic junk filtering disabled. This threshold suggests pairs only. It never approves identity changes.

Case and spacing variants, close spellings, distinct similar labels and missing area codes are covered in [the input register](input/input.xlsx) and [tests](tests/test_customers.py). Matching can miss customers with different area codes or very different names. Similar labels may describe distinct customers. Review the underlying register before confirming a link. No accuracy percentage is claimed.

## Confirmation behaviour

- A repeated identical link has the same effect as one link.
- Compatible chains resolve to their final target; the whole path appears in Audit.
- Unknown or blank IDs, self-mappings, conflicting targets and cycles reject the **entire** batch. No partial merges are applied. The command exits unsuccessfully, the report says REJECTED and Errors lists the reasons.
- A malformed register fails with a nonzero exit. `report.json` says FAIL and `outputs_current` is false; older workbook files must not be treated as this run's result.
- A rerun uses the original register and the current confirmation list. Do not use the retained register as input with the previous confirmation file: removed source IDs would then be unknown.

## Outputs

| File or tab | Meaning |
|---|---|
| result.xlsx / Customers | One row per retained ID; its chosen target's details. |
| result.xlsx / Proposed; proposed.csv | Suggestions, score, reason and review status. |
| result.xlsx / Confirmed | Submitted source-to-target links, including invalid links when rejected. |
| result.xlsx / Audit; audit.csv | Every source ID, final retained ID, confirmation path, reason, original row and original fields. |
| result.xlsx / Errors; confirmation-errors.csv | Invalid confirmation rows and reasons. |
| before.xlsx | Register as read, before identity changes. |
| checked.xlsx | LibreOffice's reopened/exported workbook, compared against result.xlsx. |
| confirmed-template.csv | Empty editable confirmation list; never populated automatically. |
| report.json; evidence.md | Counts, status, conservation check, hashes, timing and readable results. |
| runtime.json; test-results.json | Software versions, checks, limitations and test outcomes. |

CSV exports prefix text beginning with spreadsheet expression characters to keep it inert when opened. Use original customer IDs from the workbook for approvals, not escaped CSV representations. Workbooks store such text as literal strings. No fields from removed sources disappear: they remain in Audit. No rows are excluded by the merge workflow.

The view workbooks and page images show compact extracts of the same output, produced by [make_views.py](scripts/make_views.py). They are presentation evidence for the bundled register, not a separate customer list.
