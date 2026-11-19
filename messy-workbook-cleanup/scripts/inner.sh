#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL — cleanup or validation failed" >&2' ERR
python3 scripts/verify_run.py "$@"
