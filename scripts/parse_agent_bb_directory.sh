#!/usr/bin/env bash
#
# Wrapper for pe-sub-jobs/scripts/parse_agent_bb_directory.py — batch-processes Agent BB Excel
# files organized by Agent Bank subdirectories. For Linux/macOS terminals. Resolves its own
# location so it runs from any working directory, and forwards every argument through unchanged.
#
#   chmod +x pe-sub-jobs/scripts/parse_agent_bb_directory.sh   # once
#   ./pe-sub-jobs/scripts/parse_agent_bb_directory.sh <input_directory> [options]
#   ./pe-sub-jobs/scripts/parse_agent_bb_directory.sh --help
#
# Example:
#   ./pe-sub-jobs/scripts/parse_agent_bb_directory.sh data/import/agent-bank-batch/ --verbose
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY_SCRIPT="$SCRIPT_DIR/parse_agent_bb_directory.py"

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

exec "$PY" "$PY_SCRIPT" "$@"
