#!/usr/bin/env bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
trap 'echo "FAIL: combine checks" >&2' ERR
if [[ ${1:-} == --demo ]]; then
  [[ $# == 3 ]] || { echo 'Use: inner.sh --demo INPUT EMPTY_RESULT' >&2; exit 2; }
  case_root=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd -P)
  source_root=$(cd -- "$2" && pwd -P)
  mkdir -p -- "$3"
  result_root=$(cd -- "$3" && pwd -P)
  if [[ $source_root == "$result_root" || $source_root == "$result_root/"* || $result_root == "$source_root/"* ]]; then
    echo 'Input and output must be separate, non-nested folders' >&2; exit 2
  fi
  [[ -z $(find "$result_root" -mindepth 1 -maxdepth 1 -print -quit) ]] || {
    echo 'Demo results folder must be empty' >&2; exit 2;
  }
  # Copy executable sources only. Published results never seed a verification run.
  cp -R -- "$case_root/scripts" "$case_root/vba" "$case_root/tests" "$result_root/"
  cp -- "$case_root/run.sh" "$result_root/"
  mkdir -- "$result_root/input" "$result_root/output" "$result_root/evidence"
  ln -s -- "$source_root" "$result_root/input/files"
  cd -- "$result_root"
  shift 3
fi
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
