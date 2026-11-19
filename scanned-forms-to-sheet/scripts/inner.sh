#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL: scanned forms checks" >&2' ERR
input_dir=input
output_dir=output
benchmark=yes
while (($#)); do
  case "$1" in
    --input) input_dir="$2"; benchmark=no; shift 2;;
    --output) output_dir="$2"; shift 2;;
    *) echo 'Unknown argument' >&2; exit 2;;
  esac
done
args=(--input "$input_dir" --output "$output_dir")
if [[ "$benchmark" == yes ]]; then args+=(--benchmark); fi
python3 scripts/execute.py "${args[@]}"
