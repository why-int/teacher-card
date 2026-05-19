@echo off
setlocal

if not exist .venv (
    py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

set ICON_ARG=
if exist assets\app_icon.ico (
    set ICON_ARG=--icon assets\app_icon.ico
) else (
    if exist assets\app_icon.png (
        set ICON_ARG=--icon assets\app_icon.png
    )
)

set DATA_ARGS=--add-data "assets;assets"
if exist "Пример.xlsx" (
    set DATA_ARGS=%DATA_ARGS% --add-data "Пример.xlsx;templates"
)
if exist "Шаблон и пример.xlsm" (
    set DATA_ARGS=%DATA_ARGS% --add-data "Шаблон и пример.xlsm;templates"
)

pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name ExcelDataViewer ^
  %DATA_ARGS% ^
  %ICON_ARG% ^
  main.py

echo.
echo Build complete.
echo EXE: dist\ExcelDataViewer.exe
