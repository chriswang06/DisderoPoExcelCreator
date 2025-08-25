#!/bin/bash

# DisderoPoExcelCreator - Complete Installation Script
# This script installs all dependencies and sets up the application

set -e  # Exit on error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}[✓]${NC} $1"
}

print_error() {
    echo -e "${RED}[✗]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[!]${NC} $1"
}

# Function to check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Function to detect OS
detect_os() {
    if [[ "$OSTYPE" == "linux-gnu"* ]]; then
        if [ -f /etc/os-release ]; then
            . /etc/os-release
            OS=$ID
            VER=$VERSION_ID
        fi
    elif [[ "$OSTYPE" == "darwin"* ]]; then
        OS="macos"
    else
        print_error "Unsupported operating system: $OSTYPE"
        exit 1
    fi
}

# Main installation
main() {
    echo "======================================"
    echo "DisderoPoExcelCreator - Installation Script"
    echo "======================================"
    echo ""

    # Detect OS
    detect_os
    echo "Detected OS: $OS"
    echo ""

    # Check for root/sudo (needed for system packages)
    if [[ $EUID -eq 0 ]]; then
        print_warning "Running as root"
    else
        print_status "Running as user: $USER"
        print_warning "You may be prompted for sudo password"
    fi
    echo ""

    # Update package manager
    print_status "Updating package manager..."
    case $OS in
        ubuntu|debian)
            sudo apt-get update -qq
            ;;
        fedora|rhel|centos)
            sudo dnf check-update -q || true
            ;;
        arch|manjaro)
            sudo pacman -Sy --noconfirm
            ;;
        macos)
            if ! command_exists brew; then
                print_error "Homebrew not installed. Please install from https://brew.sh"
                exit 1
            fi
            brew update
            ;;
    esac

    # Install Python 3 if not present
    print_status "Checking Python..."
    if ! command_exists python3; then
        print_warning "Python 3 not found. Installing..."
        case $OS in
            ubuntu|debian)
                sudo apt-get install -y python3 python3-pip python3-venv
                ;;
            fedora|rhel|centos)
                sudo dnf install -y python3 python3-pip
                ;;
            arch|manjaro)
                sudo pacman -S --noconfirm python python-pip
                ;;
            macos)
                brew install python3
                ;;
        esac
    else
        PYTHON_VERSION=$(python3 --version | cut -d' ' -f2)
        print_status "Python $PYTHON_VERSION found"
    fi

    # Install tkinter
    print_status "Installing tkinter..."
    case $OS in
        ubuntu|debian)
            sudo apt-get install -y python3-tk
            ;;
        fedora|rhel|centos)
            sudo dnf install -y python3-tkinter
            ;;
        arch|manjaro)
            sudo pacman -S --noconfirm tk
            ;;
        macos)
            # tkinter comes with Python on macOS
            print_status "tkinter included with Python on macOS"
            ;;
    esac

    # Install Tesseract OCR
    print_status "Installing Tesseract OCR..."
    if ! command_exists tesseract; then
        case $OS in
            ubuntu|debian)
                sudo apt-get install -y tesseract-ocr
                ;;
            fedora|rhel|centos)
                sudo dnf install -y tesseract
                ;;
            arch|manjaro)
                sudo pacman -S --noconfirm tesseract
                ;;
            macos)
                brew install tesseract
                ;;
        esac
    else
        TESSERACT_VERSION=$(tesseract --version | head -n1)
        print_status "Tesseract already installed: $TESSERACT_VERSION"
    fi

    # Install Poppler utilities
    print_status "Installing Poppler utilities..."
    if ! command_exists pdftoppm; then
        case $OS in
            ubuntu|debian)
                sudo apt-get install -y poppler-utils
                ;;
            fedora|rhel|centos)
                sudo dnf install -y poppler-utils
                ;;
            arch|manjaro)
                sudo pacman -S --noconfirm poppler
                ;;
            macos)
                brew install poppler
                ;;
        esac
    else
        print_status "Poppler already installed"
    fi

    # Install development tools (needed for some Python packages)
    print_status "Installing development tools..."
    case $OS in
        ubuntu|debian)
            sudo apt-get install -y build-essential python3-dev
            ;;
        fedora|rhel|centos)
            sudo dnf groupinstall -y "Development Tools"
            sudo dnf install -y python3-devel
            ;;
        arch|manjaro)
            sudo pacman -S --noconfirm base-devel
            ;;
        macos)
            xcode-select --install 2>/dev/null || true
            ;;
    esac

    # Clone or download the project (if not already present)
    if [ ! -f "main.py" ] && [ ! -f "gui_app.py" ]; then
        print_warning "Project files not found in current directory"
        read -p "Enter project directory path (or press Enter to skip): " PROJECT_DIR
        if [ -n "$PROJECT_DIR" ]; then
            cd "$PROJECT_DIR" || exit 1
        fi
    fi

    # Create virtual environment
    print_status "Creating Python virtual environment..."
    if [ -d "venv" ]; then
        print_warning "Virtual environment already exists. Removing old one..."
        rm -rf venv
    fi
    python3 -m venv venv

    # Activate virtual environment
    print_status "Activating virtual environment..."
    source venv/bin/activate

    # Upgrade pip
    print_status "Upgrading pip..."
    pip install --upgrade pip --quiet

    # Install Python dependencies
    print_status "Installing Python dependencies..."
    if [ -f "requirements.txt" ]; then
        pip install -r requirements.txt --quiet
        print_status "All Python packages installed"
    else
        print_warning "requirements.txt not found. Installing packages manually..."
        pip install --quiet \
            opencv-python \
            Pillow \
            pytesseract \
            pandas \
            numpy \
            pdf2image \
            XlsxWriter \
            openpyxl \
            pyinstaller
    fi

    # Test installations
    echo ""
    print_status "Testing installations..."

    # Test Python modules
    python3 -c "import tkinter" && print_status "tkinter OK" || print_error "tkinter FAILED"
    python3 -c "import cv2" && print_status "OpenCV OK" || print_error "OpenCV FAILED"
    python3 -c "import PIL" && print_status "Pillow OK" || print_error "Pillow FAILED"
    python3 -c "import pytesseract" && print_status "pytesseract OK" || print_error "pytesseract FAILED"
    python3 -c "import pandas" && print_status "pandas OK" || print_error "pandas FAILED"
    python3 -c "import pdf2image" && print_status "pdf2image OK" || print_error "pdf2image FAILED"

    # Test system commands
    command_exists tesseract && print_status "Tesseract command OK" || print_error "Tesseract command FAILED"
    command_exists pdftoppm && print_status "Poppler command OK" || print_error "Poppler command FAILED"

    # Create desktop shortcut (optional)
    echo ""
    read -p "Create desktop shortcut? (y/n): " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        create_desktop_shortcut
    fi

    # Build executable (optional)
    echo ""
    read -p "Build standalone executable? (y/n): " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        build_executable
    fi

    echo ""
    echo "======================================"
    print_status "Installation complete!"
    echo "======================================"
    echo ""
    echo "To run the application:"
    echo "  GUI version:  python3 gui_app.py"
    echo "  CLI version:  python3 main.py <pdf_file>"
    echo ""
    echo "To activate the virtual environment in the future:"
    echo "  source venv/bin/activate"
    echo ""
}

# Function to create desktop shortcut
create_desktop_shortcut() {
    print_status "Creating desktop shortcut..."

    DESKTOP_DIR="$HOME/Desktop"
    if [ ! -d "$DESKTOP_DIR" ]; then
        DESKTOP_DIR="$HOME"
    fi

    cat > "$DESKTOP_DIR/POProcessor.desktop" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=DisderoPoExcelCreator
Comment=Process Purchase Orders
Exec=$(pwd)/venv/bin/python $(pwd)/gui_app.py
Icon=application-pdf
Terminal=false
Categories=Office;
EOF

    chmod +x "$DESKTOP_DIR/POProcessor.desktop"
    print_status "Desktop shortcut created at $DESKTOP_DIR/POProcessor.desktop"
}

# Function to build executable
build_executable() {
    print_status "Building standalone executable..."

    if [ -f "POProcessor.spec" ]; then
        pyinstaller POProcessor.spec
    else
        pyinstaller --onefile --windowed --name=POProcessor gui_app.py
    fi

    if [ -f "dist/POProcessor" ]; then
        print_status "Executable built successfully at dist/POProcessor"
        chmod +x dist/POProcessor

        # Copy master file if exists
        if [ -f "productslist.xlsx" ]; then
            cp productslist.xlsx dist/
            print_status "Copied productslist.xlsx to dist/"
        fi
    else
        print_error "Failed to build executable"
    fi
}

# Run main installation
main "$@"