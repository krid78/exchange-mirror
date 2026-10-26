@echo off
setlocal
pushd "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Create .venv and install requirements-importer-build.txt first.
    popd
    exit /b 1
)

echo === Building ICS Calendar Importer ===
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onedir --windowed ^
    --name ics-importer ^
    --distpath dist --workpath build\ics-importer --specpath build\ics-importer ^
    --collect-submodules icalendar.timezone --copy-metadata icalendar ^
    --collect-data tzdata ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone ^
    ics_importer.py
if errorlevel 1 (
    echo IMPORTER BUILD FAILED
    popd
    exit /b 1
)
echo Output: dist\ics-importer\ics-importer.exe
echo Distribute the entire dist\ics-importer folder including _internal.
popd
exit /b 0
