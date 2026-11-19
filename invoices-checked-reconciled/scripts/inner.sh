#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL: invoice checks stopped" >&2' ERR
python3 scripts/run_pipeline.py "$@"
