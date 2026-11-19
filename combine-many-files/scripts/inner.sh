#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL: combine checks" >&2' ERR
if (($#)); then
  python3 scripts/combine.py "$@"
  python3 scripts/check_output.py "$@"
else
  python3 scripts/combine.py
  /opt/vba/bin/python scripts/package_vba.py
  python3 scripts/run_vba.py
  python3 -m pytest -q -p no:cacheprovider tests
fi
echo 'PASS: combine checks'
