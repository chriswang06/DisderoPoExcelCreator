@echo off
setlocal enabledelayedexpansion

:: PO Processor - Windows Installation Script
:: This script installs all dependencies and sets up the application

echo ======================================
echo PO Processor - Windows Installation
echo ======================================
echo.

:: Check for admin rights (needed for some installations)
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges
) else (
    echo [!] Running without admin privileges
    echo     Some installations may require manual intervention
)
echo.

:: Check for Python
echo Checking for Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [X] Python not found!
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
) else (
    for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
    echo [OK] Python !PYTHON_VERSION! found
)

:: Check for pip
echo Checking for pip...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] pip not found, installing...
    python -m ensurepip --upgrade
)

:: Check for Tesseract
echo.
echo Checking for Tesseract OCR...
tesseract --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [X] Tesseract not found!
    echo.
    echo Please install Tesseract:
    echo 1. Download from: https://github.com/UB-Mannheim/tesseract/wiki
    echo 2. Run the installer
    echo 3. Add Tesseract to your PATH or note the installation directory
    echo    Default: C:\Program Files\Tesseract-OCR
    echo.
    set /p TESSERACT_PATH="Enter Tesseract installation path (or press Enter to skip): "
    if not "!TESSERACT_PATH!"=="" (
        setx PATH "%PATH%;!TESSERACT_PATH!" >nul 2>&1
        echo Added !TESSERACT_PATH! to PATH
    )
) else (
    echo [OK] Tesseract found
)

:: Check for Poppler
echo Checking for Poppler utilities...
pdftoppm -h >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Poppler not found
    echo.
    echo Poppler is required for PDF processing.
    echo.
    echo Automatic download option:
    set /p INSTALL_POPPLER="Download and install Poppler? (y/n): "
    if /i "!INSTALL_POPPLER!"=="y" (
        call :install_poppler
    ) else (
        echo.
        echo Manual installation:
        echo 1. Download from: https://github.com/oschwartz10612/poppler-windows/releases/
        echo 2. Extract to C:\poppler or another directory
        echo 3. Add the 'bin' folder to your PATH
        echo.
    )
) else (
    echo [OK] Poppler found
)

:: Create virtual environment
echo.
echo Creating Python virtual environment...
if exist venv (
    echo [!] Virtual environment already exists, removing...
    rmdir /s /q venv
)
python -m venv venv

:: Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

:: Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip --quiet

:: Install Python packages
echo.
echo Installing Python packages...
if exist requirements.txt (
    pip install -r requirements.txt --quiet
    echo [OK] All Python packages installed
) else (
    echo [!] requirements.txt not found, installing packages manually...
    pip install --quiet ^
        opencv-python ^
        Pillow ^
        pytesseract ^
        pandas ^
        numpy ^
        pdf2image ^
        XlsxWriter ^
        openpyxl ^
        pyinstaller
)

:: Test installations
echo.
echo Testing installations...
python -c "import tkinter" 2>nul && echo [OK] tkinter || echo [X] tkinter FAILED
python -c "import cv2" 2>nul && echo [OK] OpenCV || echo [X] OpenCV FAILED
python -c "import PIL" 2>nul && echo [OK] Pillow || echo [X] Pillow FAILED
python -c "import pytesseract" 2>nul && echo [OK] pytesseract || echo [X] pytesseract FAILED
python -c "import pandas" 2>nul && echo [OK] pandas || echo [X] pandas FAILED
python -c "import pdf2image" 2>nul && echo [OK] pdf2image || echo [X] pdf2image FAILED

:: Create run.bat for easy launching
echo.
echo Creating run.bat launcher...
(
echo @echo off
echo call venv\Scripts\activate.bat
echo python gui_app.py
echo pause
) > run.bat
echo [OK] Created run.bat

:: Create desktop shortcut (optional)
echo.
set /p CREATE_SHORTCUT="Create desktop shortcut? (y/n): "
if /i "%CREATE_SHORTCUT%"=="y" (
    call :create_shortcut
)

:: Build executable (optional)
echo.
set /p BUILD_EXE="Build standalone executable? (y/n): "
if /i "%BUILD_EXE%"=="y" (
    call :build_executable
)

echo.
echo ======================================
echo [OK] Installation complete!
echo ======================================
echo.
echo To run the application:
echo   GUI: Double-click run.bat
echo   CLI: python main.py [pdf_file]
echo.
echo To activate virtual environment manually:
echo   venv\Scripts\activate.bat
echo.
pause
exit /b 0

:: Function to install Poppler
:install_poppler
echo Downloading Poppler...
powershell -Command "& {Invoke-WebRequest -Uri 'https://github.com/oschwartz10612/poppler-windows/releases/latest/download/Release-23.08.0-0.zip' -OutFile 'poppler.zip'}"
if exist poppler.zip (
    echo Extracting Poppler...
    powershell -Command "& {Expand-Archive -Path 'poppler.zip' -DestinationPath 'C:\' -Force}"
    del poppler.zip
    echo Adding Poppler to PATH...
    setx PATH "%PATH%;C:\poppler-23.08.0\Library\bin" >nul 2>&1
    echo [OK] Poppler installed to C:\poppler-23.08.0
    echo [!] Please restart the command prompt for PATH changes to take effect
) else (
    echo [X] Failed to download Poppler
)
goto :eof

:: Function to create desktop shortcut
:create_shortcut
echo Creating desktop shortcut...
set DESKTOP=%USERPROFILE%\Desktop
set SCRIPT_PATH=%~dp0

powershell -Command "& {$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%DESKTOP%\PO Processor.lnk'); $Shortcut.TargetPath = '%SCRIPT_PATH%run.bat'; $Shortcut.WorkingDirectory = '%SCRIPT_PATH%'; $Shortcut.IconLocation = 'shell32.dll,21'; $Shortcut.Save()}"
echo [OK] Desktop shortcut created
goto :eof

:: Function to build executable
:build_executable
echo Building standalone executable...
if exist POProcessor.spec (
    pyinstaller POProcessor.spec
) else (
    pyinstaller --onefile --windowed --name=POProcessor gui_app.py
)
if exist dist\POProcessor.exe (
    echo [OK] Executable built at dist\POProcessor.exe
    if exist productslist.xlsx (
        copy productslist.xlsx dist\
        echo [OK] Copied productslist.xlsx to dist\
    )
) else (
    echo [X] Failed to build executable
)
goto :eof