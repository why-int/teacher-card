#!/usr/bin/env bash
set -euo pipefail

if [[ ! -d ".venv" ]]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt

ICON_ARG=""
if [[ -f "assets/app_icon.icns" ]]; then
  ICON_ARG="assets/app_icon.icns"
elif [[ -f "assets/app_icon.png" ]]; then
  ICON_ARG="assets/app_icon.png"
fi

if [[ -n "${ICON_ARG}" ]]; then
  pyinstaller \
    --noconfirm \
    --clean \
    --windowed \
    --name "ExcelDataViewer" \
    --add-data "assets:assets" \
    --icon "${ICON_ARG}" \
    main.py
else
  pyinstaller \
    --noconfirm \
    --clean \
    --windowed \
    --name "ExcelDataViewer" \
    --add-data "assets:assets" \
    main.py
fi

echo ""
echo "Build complete."
echo "APP: dist/ExcelDataViewer.app"
