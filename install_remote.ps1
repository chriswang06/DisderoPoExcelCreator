# DisderoPoExcelCreator - Windows Remote Installation Script
# Run with: iwr -useb https://raw.githubusercontent.com/chriswang06/DisderoPoExcelCreator/main/install_remote.ps1 | iex

# Check for admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

Write-Host "=====================================" -ForegroundColor Blue
Write-Host "   DisderoPoExcelCreator - Quick Installer   " -ForegroundColor Blue
Write-Host "=====================================" -ForegroundColor Blue
Write-Host ""

# GitHub repository
$repoUrl = "https://github.com/chriswang06/DisderoPoExcelCreator.git"
$defaultInstallDir = "$env:USERPROFILE\DisderoPoExcelCreator"

# Check Git
Write-Host "Checking for Git..." -ForegroundColor Green
$gitInstalled = Get-Command git -ErrorAction SilentlyContinue

if (-not $gitInstalled) {
    Write-Host "Git not found. Installing..." -ForegroundColor Yellow

    # Download and install Git
    $gitInstaller = "$env:TEMP\git-installer.exe"
    Write-Host "Downloading Git installer..."
    Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/download/v2.43.0.windows.1/Git-2.43.0-64-bit.exe" -OutFile $gitInstaller

    Write-Host "Installing Git (this may take a moment)..."
    Start-Process -FilePath $gitInstaller -ArgumentList "/VERYSILENT", "/NORESTART" -Wait
    Remove-Item $gitInstaller

    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# Check Python
Write-Host "Checking for Python..." -ForegroundColor Green
$pythonInstalled = Get-Command python -ErrorAction SilentlyContinue

if (-not $pythonInstalled) {
    Write-Host "Python not found!" -ForegroundColor Red
    Write-Host "Please install Python from: https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Make sure to check 'Add Python to PATH' during installation" -ForegroundColor Yellow

    $installPython = Read-Host "Would you like to download Python installer now? (y/n)"
    if ($installPython -eq 'y') {
        Start-Process "https://www.python.org/downloads/"
    }
    Write-Host "Please run this script again after installing Python" -ForegroundColor Yellow
    pause
    exit 1
}

# Get installation directory
Write-Host ""
Write-Host "Where would you like to install DisderoPoExcelCreator?" -ForegroundColor Green
Write-Host "Default: $defaultInstallDir" -ForegroundColor Yellow
$installDir = Read-Host "Press Enter for default or enter path"

if ([string]::IsNullOrWhiteSpace($installDir)) {
    $installDir = $defaultInstallDir
}

# Check if directory exists
if (Test-Path $installDir) {
    Write-Host "Directory $installDir already exists" -ForegroundColor Yellow
    $response = Read-Host "Remove existing installation? (y/n)"
    if ($response -eq 'y') {
        Remove-Item -Recurse -Force $installDir
    } else {
        Write-Host "Installation cancelled" -ForegroundColor Red
        pause
        exit 1
    }
}

# Clone repository
Write-Host ""
Write-Host "Cloning repository..." -ForegroundColor Green
git clone $repoUrl $installDir

# Change to installation directory
Set-Location $installDir

# Check for Tesseract
Write-Host ""
Write-Host "Checking for Tesseract OCR..." -ForegroundColor Green
$tesseractInstalled = Get-Command tesseract -ErrorAction SilentlyContinue

if (-not $tesseractInstalled) {
    Write-Host "Tesseract not found!" -ForegroundColor Yellow
    Write-Host "Downloading Tesseract installer..." -ForegroundColor Green

    $tesseractUrl = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3.20231005/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    $tesseractInstaller = "$env:TEMP\tesseract-installer.exe"

    try {
        Invoke-WebRequest -Uri $tesseractUrl -OutFile $tesseractInstaller
        Write-Host "Please install Tesseract when the installer opens" -ForegroundColor Yellow
        Write-Host "IMPORTANT: Remember the installation path!" -ForegroundColor Red
        Start-Process -FilePath $tesseractInstaller -Wait
        Remove-Item $tesseractInstaller

        # Add to PATH
        $tesseractPath = "C:\Program Files\Tesseract-OCR"
        if (Test-Path $tesseractPath) {
            [Environment]::SetEnvironmentVariable("Path", $env:Path + ";$tesseractPath", [EnvironmentVariableTarget]::User)
            $env:Path += ";$tesseractPath"
        }
    } catch {
        Write-Host "Failed to download Tesseract" -ForegroundColor Red
        Write-Host "Please install manually from: https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Yellow
    }
}

# Check for Poppler
Write-Host ""
Write-Host "Checking for Poppler..." -ForegroundColor Green
$popplerInstalled = Get-Command pdftoppm -ErrorAction SilentlyContinue

if (-not $popplerInstalled) {
    Write-Host "Poppler not found. Installing..." -ForegroundColor Yellow

    $popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v23.08.0-0/Release-23.08.0-0.zip"
    $popplerZip = "$env:TEMP\poppler.zip"
    $popplerDir = "C:\poppler"

    try {
        Write-Host "Downloading Poppler..."
        Invoke-WebRequest -Uri $popplerUrl -OutFile $popplerZip

        Write-Host "Extracting Poppler..."
        Expand-Archive -Path $popplerZip -DestinationPath "C:\" -Force
        Remove-Item $popplerZip

        # Add to PATH
        $popplerBin = "C:\poppler-23.08.0\Library\bin"
        if (Test-Path $popplerBin) {
            [Environment]::SetEnvironmentVariable("Path", $env:Path + ";$popplerBin", [EnvironmentVariableTarget]::User)
            $env:Path += ";$popplerBin"
            Write-Host "Poppler installed successfully" -ForegroundColor Green
        }
    } catch {
        Write-Host "Failed to install Poppler automatically" -ForegroundColor Red
        Write-Host "Please install manually from: https://github.com/oschwartz10612/poppler-windows/releases/" -ForegroundColor Yellow
    }
}

# Run the local install.bat
Write-Host ""
if (Test-Path "install.bat") {
    Write-Host "Running installation script..." -ForegroundColor Green
    & cmd.exe /c "install.bat"
} else {
    # Fallback installation
    Write-Host "install.bat not found, performing basic setup..." -ForegroundColor Yellow

    # Create virtual environment
    python -m venv venv

    # Activate and install packages
    & ".\venv\Scripts\Activate.ps1"
    python -m pip install --upgrade pip

    if (Test-Path "requirements.txt") {
        pip install -r requirements.txt
    }

    # Create run.bat
    @"
@echo off
call venv\Scripts\activate.bat
python gui_app.py
pause
"@ | Out-File -FilePath "run.bat" -Encoding ASCII
}

# Create desktop shortcut
Write-Host ""
$createShortcut = Read-Host "Create desktop shortcut? (y/n)"
if ($createShortcut -eq 'y') {
    $desktop = [Environment]::GetFolderPath("Desktop")
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$desktop\DisderoPoExcelCreator.lnk")
    $Shortcut.TargetPath = "$installDir\run.bat"
    $Shortcut.WorkingDirectory = $installDir
    $Shortcut.IconLocation = "shell32.dll,21"
    $Shortcut.Save()
    Write-Host "Desktop shortcut created" -ForegroundColor Green
}

# Create Start Menu entry
$startMenu = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$startMenu\DisderoPoExcelCreator.lnk")
$Shortcut.TargetPath = "$installDir\run.bat"
$Shortcut.WorkingDirectory = $installDir
$Shortcut.IconLocation = "shell32.dll,21"
$Shortcut.Save()

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "    Installation Complete!           " -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installation directory: $installDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "To run DisderoPoExcelCreator:" -ForegroundColor Blue
Write-Host "  1. Double-click the desktop shortcut" -ForegroundColor White
Write-Host "  2. Or run: $installDir\run.bat" -ForegroundColor White
Write-Host "  3. Or find it in your Start Menu" -ForegroundColor White
Write-Host ""
Write-Host "To update in the future:" -ForegroundColor Blue
Write-Host "  cd $installDir" -ForegroundColor White
Write-Host "  git pull" -ForegroundColor White
Write-Host ""
pause