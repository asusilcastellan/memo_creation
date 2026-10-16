#!/usr/bin/env bash
# macOS / Linux build (same as build.bat)
set -e
pyinstaller memo_app.spec --noconfirm --clean
cp -R resources dist/CastellanMemo/
cp .env.example dist/CastellanMemo/
[ -f .env ] && cp .env dist/CastellanMemo/
echo "Built: dist/CastellanMemo/CastellanMemo"
