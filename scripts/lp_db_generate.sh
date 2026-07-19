#!/usr/bin/env bash
#
# Wrapper for pe-sub-jobs/scripts/lp_db_generate.py — regenerates the simulated, chaos-degraded
# LP DB Export XLSX in data/import/ — for Linux/macOS terminals. Resolves its own location so it
# runs from any working directory.
#
# There are NO arguments: tune the constants near the top of lp_db_generate.py (SEED,
# CHAOS_ENABLED, CHAOS_SEED, TARGET_ROWS, ...), then run this wrapper. Re-run
# lp_db_extract.sh afterwards to produce the seed CSVs.
#
#   chmod +x pe-sub-jobs/scripts/lp_db_generate.sh   # once
#   ./pe-sub-jobs/scripts/lp_db_generate.sh          # or: bash pe-sub-jobs/scripts/lp_db_generate.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY_SCRIPT="$SCRIPT_DIR/lp_db_generate.py"

if [[ ! -f "$PY_SCRIPT" ]]; then
  echo "error: cannot find $PY_SCRIPT" >&2
  exit 1
fi

# Prefer python3, fall back to python.
if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo "error: python3 (or python) not found on PATH" >&2
  exit 1
fi

# openpyxl is the only third-party dependency.
if ! "$PY" -c "import openpyxl" >/dev/null 2>&1; then
  echo "error: the 'openpyxl' package is required. Install it with:" >&2
  echo "  $PY -m pip install openpyxl" >&2
  exit 1
fi

exec "$PY" "$PY_SCRIPT"
