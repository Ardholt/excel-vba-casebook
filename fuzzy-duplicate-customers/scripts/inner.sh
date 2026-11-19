#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL: customer checks did not complete" >&2' ERR
INPUT=input
OUTPUT=output
while (($#)); do
  case "$1" in
    --input) INPUT=$2; shift 2 ;;
    --output) OUTPUT=$2; shift 2 ;;
    *) echo "Unknown option: $1" >&2; exit 2 ;;
  esac
done
python3 scripts/customers.py --input "$INPUT" --output "$OUTPUT"
xlrecalc "$OUTPUT/result.xlsx" "$OUTPUT/checked.xlsx"
python3 scripts/check_run.py --input "$INPUT" --output "$OUTPUT"
CASE_TEST_REPORT="$OUTPUT/test-results.json" python3 -m pytest -q -p no:cacheprovider tests
echo "PASS: every source ID accounted for; only confirmed identities changed"
