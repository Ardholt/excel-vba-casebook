# Run the scanned-form workflow

Download or clone the complete repository. With Docker installed and its engine running, open a terminal in `scanned-forms-to-sheet` and run `./run.sh`. The [runner](run.sh) builds the [shared container](../docker/Dockerfile), regenerates `output/` and executes the business tests. No host Python or office application is needed. The first image build needs package-source access; processing runs without network access.

For another batch:

```bash
./run.sh --input "./my forms" --output "./my results"
```

Keep PNGs directly inside the input directory's `forms/` subfolder. Both paths may contain spaces. Sources are mounted read-only. Use separate input/output directories: the named result files in the output directory are replaced on each run.

Use the printed [layout specification](LAYOUT.md). It defines image dimensions, field rectangles, accepted values and review keys. The reader supports the documented upright layout, not arbitrary invoices or handwriting. Every supplied PNG is represented, including unreadable files. Empty batches and invalid review rows fail explicitly.

Optional `review.csv` has named columns `form_id,field,value` in any order. Use the queue's review key and the exact field labels. Unknown targets, repeated targets and invalid values are reported. The raw reading stays unchanged; accepted confirmations appear in a separate result and audit.

| Workbook | Tabs |
|---|---|
| `input-inventory.xlsx` | `Scans`: image names, dimensions and hashes |
| `extracted.xlsx` | `Data`: raw accepted values; `Unclear entries`: all originally flagged fields |
| `reviewed.xlsx` | `Data`: values after supplied confirmations; `Still unresolved`: remaining blanks; `Review audit`: original and confirmed values |

`fields.csv` retains candidates, confidence, accepted values and source rectangles. `unclear-entries.csv` preserves the original queue; `unresolved.csv` contains the remaining queue. `report.json` reports counts, checks and extraction time.

Replacement batches are not scored against the bundled key. To score another batch, run `python3 scripts/evaluate.py OUTPUT KEY.csv` inside the same container after extraction. The key must have `image,field,value` and cover every field exactly once. The evaluator cannot fill output values. The extractor never loads the key or the form generator.

Accuracy in the public report applies only to the supplied image set. Supplied review values are not evidence of an independent human review. Confidence is a routing heuristic, not a probability of correctness.

[Dependency pins and asset licences](DEPENDENCIES.md) · [executed versions and checks](evidence/runtime.json).
