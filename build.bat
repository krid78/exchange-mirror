@echo off
echo === Building exchange-mirror ===

echo.
echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo DEPENDENCY INSTALL FAILED
    exit /b 1
)

echo.
echo Running PyInstaller...
python -m PyInstaller -y --onedir --windowed --name mirror ^
    --exclude-module config ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import win32api ^
    --hidden-import win32timezone ^
    --hidden-import pystray ^
    --hidden-import pystray._win32 ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    --hidden-import PIL.ImageDraw ^
    --collect-all tkcalendar ^
    --collect-all babel ^
    tray.py

if errorlevel 1 (
    echo.
    echo BUILD FAILED
    exit /b 1
)

echo.
echo Copying config.py into dist\mirror\ ...
copy /Y config.py dist\mirror\config.py

echo.
echo Building installer...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if errorlevel 1 (
    echo.
    echo INSTALLER BUILD FAILED
    echo Make sure Inno Setup 6 is installed: choco install innosetup -y
    exit /b 1
)

echo.
echo === Build complete ===
echo Output: dist\Exchange_Mirror_Setup.exe
