@echo off
echo ============================================
echo   Building Voice Typer .exe
echo ============================================
echo.

:: Check that PyInstaller is installed
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [1/4] Installing PyInstaller...
    pip install pyinstaller
) else (
    echo [1/4] PyInstaller found.
)

echo [2/4] Generating icon.ico...
python voice_typer.py --make-icon
if errorlevel 1 (
    echo.
    echo [WARNING] Icon generation failed, building without icon.
    set ICON_FLAG=
) else (
    echo       icon.ico generated.
    set ICON_FLAG=--icon icon.ico
)

echo [3/4] Building executable...
echo.

pyinstaller ^
    --onefile ^
    --noconsole ^
    --collect-data customtkinter ^
    --hidden-import pynput.keyboard._win32 ^
    --hidden-import pynput.mouse._win32 ^
    --name VoiceTyper ^
    --clean ^
    %ICON_FLAG% ^
    voice_typer.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo ============================================
echo   BUILD COMPLETE
echo   Output: dist\VoiceTyper.exe
echo ============================================
echo.
echo You can now copy dist\VoiceTyper.exe anywhere
echo and run it — no Python needed.
echo.
pause
