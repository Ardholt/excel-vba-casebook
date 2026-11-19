# Proof images

These PNGs are actual LibreOffice sheet exports, cropped by the sheet renderer to their content and saved without identifying image metadata. `balance-checks.png` is the opening proof; `statement-layouts.png` shows the original field/value layouts; `trace-a-balance.png` shows the flagged ledger row and its source reference.

[Image hashes, worksheet names and workbook hashes](../evidence/images.json) tie each capture to its source. Regenerate after a successful default run with `python3 scripts/render_proof.py` inside the shared container. The test badge records the captured container test result.
