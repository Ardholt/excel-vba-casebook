#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL: dashboard refresh or checks failed" >&2' ERR
output_dir=output
for ((i=1; i<=$#; i++)); do
  if [[ "${!i}" == --output ]]; then j=$((i+1)); output_dir="${!j}"; fi
done
python3 scripts/run_reference.py "$@"
/opt/vba/bin/python scripts/package_macro.py "$@"
python3 scripts/run_vba.py "$@"
/opt/vba/bin/python scripts/package_macro.py "$@" --verify
python3 scripts/test_run.py --output "$output_dir"
echo 'PASS: dashboard reference, three LibreOffice button refreshes and business tests'
