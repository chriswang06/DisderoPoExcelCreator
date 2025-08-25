#!/bin/bash
# build.sh - Build script for Linux/macOS

echo "PO Processor - Build Script"
echo "=========================="
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install/update dependencies
echo "Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

# Clean previous builds
echo "Cleaning previous builds..."
rm -rf build dist

# Build the executable
echo "Building executable..."
pyinstaller POProcessor.spec

# Check if build was successful
if [ -f "dist/POProcessor" ] || [ -f "dist/POProcessor.app" ]; then
    echo ""
    echo "✓ Build successful!"
    echo "Executable location: $(pwd)/dist/"

    # Copy master file if it exists
    if [ -f "productslist.xlsx" ]; then
        cp productslist.xlsx dist/
        echo "✓ Copied productslist.xlsx to dist folder"
    fi
else
    echo ""
    echo "✗ Build failed. Check the error messages above."
    exit 1
fi

echo ""
echo "Done!"

# --- Windows batch file content below ---
# Save this as build.bat for Windows:

: <<'BATCH_SCRIPT'
@echo off
echo PO Processor - Build Script
echo ==========================
echo.

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install/update dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

REM Clean previous builds
echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build the executable
echo Building executable...
pyinstaller POProcessor.spec

REM Check if build was successful
if exist "dist\POProcessor.exe" (
    echo.
    echo Build successful!
    echo Executable location: %cd%\dist\

    REM Copy master file if it exists
    if exist "productslist.xlsx" (
        copy productslist.xlsx dist\
        echo Copied productslist.xlsx to dist folder
    )
) else (
    echo.
    echo Build failed. Check the error messages above.
    exit /b 1
)

echo.
echo Done!
pause
BATCH_SCRIPT