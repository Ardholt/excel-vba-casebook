#!/usr/bin/env bash
set -euo pipefail
trap 'echo "FAIL: macro pack checks" >&2' ERR
export PYTHONDONTWRITEBYTECODE=1
python3 scripts/reference.py "$@"
in=input; out=output
while (($#)); do case "$1" in --input) in=$2; shift 2;; --output) out=$2; shift 2;; *) exit 2;; esac; done
python3 scripts/check_output.py "$in" "$out"
python3 -m pytest -q -p no:cacheprovider tests
printf 'PASS: four reference actions checked\n'
