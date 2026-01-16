@echo off
chcp 65001 >nul 2>&1
echo Building desktop application...
echo.

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
)

REM Install dependencies
echo Installing dependencies...
python -m pip install -r requirements.txt

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build application using spec file
echo.
echo Building application...
pyinstaller build.spec

if errorlevel 1 (
    echo.
    echo BUILD ERROR!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo Executable file is in the 'dist' folder
echo ========================================
echo.
pause

