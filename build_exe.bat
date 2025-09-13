@echo off
setlocal enabledelayedexpansion

REM Simple build script for docx_extractor.py
REM Creates standalone executable using system Python environment

echo ====================================
echo Building docx_extractor.exe
echo ====================================

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Check if we're in the correct directory
if not exist "src\main.py" (
    echo ERROR: main.py not found in src directory
    echo Please run this script from the project root directory
    pause
    exit /b 1
)

REM Clean previous build artifacts (keep dist directory)
echo Cleaning previous build artifacts...
if exist "build" rmdir /s /q "build" >nul 2>&1
if exist "dist\docx_extractor.exe" del /q "dist\docx_extractor.exe" >nul 2>&1
if exist "*.spec" del /q "*.spec" >nul 2>&1

REM Install required packages
echo Installing required packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install required packages
    pause
    exit /b 1
)

REM Build executable with PyInstaller
echo Building executable...
pyinstaller --onefile --name "docx_extractor" --clean --windowed "src\main.py"
if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed
    pause
    exit /b 1
)

REM Check if executable was created
if exist "dist\docx_extractor.exe" (
    echo ====================================
    echo BUILD SUCCESSFUL!
    echo ====================================
    echo Executable created: dist\docx_extractor.exe
    echo File size: 
    for %%F in ("dist\docx_extractor.exe") do echo %%~zF bytes
    echo.
    echo Usage: docx_extractor.exe [docx_file_path]
    echo.
) else (
    echo ERROR: Executable not found in dist directory
    pause
    exit /b 1
)

REM Clean build artifacts but keep the exe
echo Cleaning build artifacts...
if exist "build" rmdir /s /q "build" >nul 2>&1
if exist "*.spec" del /q "*.spec" >nul 2>&1

REM Remove Python cache directories
for /d /r . %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d" >nul 2>&1
)

REM Remove .pyc files
for /r . %%f in (*.pyc) do (
    if exist "%%f" del /q "%%f" >nul 2>&1
)

echo ====================================
echo Build completed successfully!
echo Only dist\docx_extractor.exe remains
echo ====================================

pause 