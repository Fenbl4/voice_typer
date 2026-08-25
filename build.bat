@echo off
echo ============================================
echo   Building Voice Typer .exe
echo ============================================
echo.

set "PYTHON_CMD=python"
py -3.13 --version >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3.13"

:: Check that PyInstaller is installed
%PYTHON_CMD% -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [1/4] Installing PyInstaller...
    %PYTHON_CMD% -m pip install pyinstaller
) else (
    echo [1/4] PyInstaller found.
)

echo [2/4] Generating icon.ico...
%PYTHON_CMD% voice_typer.py --make-icon
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

%PYTHON_CMD% -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --collect-data customtkinter ^
    --collect-data onnx_asr ^
    --copy-metadata onnx-asr ^
    --exclude-module huggingface_hub ^
    --exclude-module fsspec ^
    --exclude-module gradio ^
    --exclude-module pandas ^
    --exclude-module scipy ^
    --exclude-module numba ^
    --exclude-module llvmlite ^
    --exclude-module cv2 ^
    --exclude-module torch ^
    --exclude-module tensorflow ^
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
