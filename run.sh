#!/bin/bash
# Script untuk menjalankan generator template X-Ray secara mudah
set -e

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
cd "$DIR"

if [ -f ".venv/bin/python" ]; then
    PYTHON_BIN=".venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="python3"
else
    echo "Python tidak ditemukan!"
    exit 1
fi

"$PYTHON_BIN" main.py "$@"

