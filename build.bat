@echo off
setlocal
cd /d "%~dp0"
echo === Building exchange-mirror ===

rem Honor an explicit compiler path, then discover user/system installations.
if not defined ISCC_EXE if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do if not defined ISCC_EXE set "ISCC_EXE=%%I"
if not exist "%ISCC_EXE%" (
    echo Inno Setup 6 compiler not found. Install it or set ISCC_EXE to ISCC.exe.
    exit /b 1
)

echo.
echo Installing dependencies...
python -m pip install -r requirements.txt
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
"%ISCC_EXE%" installer.iss
if errorlevel 1 (
    echo.
    echo INSTALLER BUILD FAILED
    echo Make sure Inno Setup 6 is installed: choco install innosetup -y
    exit /b 1
)

echo.
echo === Build complete ===
echo Output: dist\Exchange_Mirror_Setup.exe
