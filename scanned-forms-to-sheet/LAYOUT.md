# Supported form layout

Use PNG images at **1000 × 600 pixels**, upright, with the same fixed field locations as [this form](input/forms/entry-01.png). These are pixel coordinates; a DPI metadata value is not required. Do not resize, rotate or move the fields. Typed dark text on a light background is supported. Layout detection, handwriting and multipage files are outside this reader.

| Label | Accepted value | Crop: left, top, right, bottom |
|---|---|---|
| Form ID | `F` then 3–8 digits | 410, 147, 945, 214 |
| Item code | `IT-` then 3–8 digits | 410, 247, 945, 314 |
| Quantity | 1–6 digits, nonnegative whole number | 410, 347, 945, 414 |
| Amount | Nonnegative amount, dot and exactly 2 decimal places; up to 9 digits before the dot | 410, 447, 945, 514 |

Coordinates and formats come from [the executable reader](scripts/extract.py). Rectangles use the Pillow convention: left/top included, right/bottom excluded. Amounts have no currency conversion or implied currency. Quantity and amount are numeric workbook cells; identifiers remain text.

The reader discovers all files whose extension is `.png`, ignoring case, directly in `forms/`. Other formats and nested folders are not inputs. Every discovered PNG produces a row and four field records, including unreadable files. A batch without PNGs fails. Supported dimensions alone do not prove the page has the right layout: inspect a batch before use.

Each field is converted to grayscale, trimmed to the dark-pixel bounds, padded with white space and enlarged before Tesseract English reads it in single-line mode. This standardises the image scale; it does not substitute characters. A value is accepted only when every recognised word meets the confidence cutoff of **85** and the full value matches its format. No substitutions such as letter O to zero are made. Empty, low-confidence and invalid-format readings stay blank. Threshold and crop values are defined in [extract.py](scripts/extract.py).

`input/review.csv` is optional. Its required headers are `form_id,field,value` and are looked up by name. Field names are exactly the labels above. Resolve `file:filename.png` review keys by looking at the image; the review key stays stable even when the Form ID itself is corrected. A Form ID correction does not redirect other confirmations to another image. The audit retains the original key.

The scoring key is a separate CSV with `image,field,value`. It must cover every extracted field exactly once. It is read only by [evaluate.py](scripts/evaluate.py), after extraction. The [form generator](scripts/make_forms.py) is for reproducing the bundled source set and key; it is never called by `run.sh`.
