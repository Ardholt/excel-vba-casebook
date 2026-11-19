#!/usr/bin/env bash
set -euo pipefail
trap 'echo "FAIL: customer workflow did not complete" >&2' ERR
CASE_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
INPUT="$CASE_DIR/input"
OUTPUT="$CASE_DIR/output"
while (($#)); do
  case "$1" in
    --input|--output)
      (($# >= 2)) || { echo "Missing folder argument" >&2; exit 2; }
      if [[ "$1" == --input ]]; then INPUT=$2; else OUTPUT=$2; fi
      shift 2 ;;
    *) echo "Usage: ./run.sh [--input DIR --output DIR]" >&2; exit 2 ;;
  esac
done
INPUT=$(cd -- "$INPUT" && pwd -P)
mkdir -p -- "$OUTPUT"
OUTPUT=$(cd -- "$OUTPUT" && pwd -P)
if [[ "$INPUT" == "$OUTPUT" || "$INPUT/" == "$OUTPUT/"* || "$OUTPUT/" == "$INPUT/"* ]]; then
  echo "Input and output folders must be separate" >&2; exit 2
fi
docker build -t excel-vba-casebook "$CASE_DIR/../docker"
docker run --rm --user "$(id -u):$(id -g)" --network none \
  --mount "type=bind,src=$CASE_DIR,dst=/case,readonly" \
  --mount "type=bind,src=$INPUT,dst=/data-in,readonly" \
  --mount "type=bind,src=$OUTPUT,dst=/data-out" \
  --workdir /case -e HOME=/tmp -e PYTHONDONTWRITEBYTECODE=1 \
  excel-vba-casebook bash scripts/inner.sh --input /data-in --output /data-out
