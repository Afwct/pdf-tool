#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then PY=python3
elif command -v python >/dev/null 2>&1; then PY=python
else echo "ERROR: 未找到 Python"; exit 1; fi

"$PY" "$(dirname "$0")/common/install_deps.py"
