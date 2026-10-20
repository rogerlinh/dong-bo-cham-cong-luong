#!/bin/sh
# Mở tool đồng bộ chấm công lương (macOS / Linux). Nhấn Ctrl+C để tắt.
cd "$(dirname "$0")" || exit 1
PY=python3
command -v python3 >/dev/null 2>&1 || PY=python
if ! "$PY" -c "import openpyxl" >/dev/null 2>&1; then
  echo "Đang cài thư viện openpyxl lần đầu..."
  "$PY" -m pip install --user -r ../requirements.txt || exit 1
fi
exec "$PY" app_web.py
