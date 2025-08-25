#!/bin/bash

# PO Processor - Remote Installation Script
# This script can be run directly from GitHub with curl/wget

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

# GitHub repository URL (update this with your actual username)
REPO_URL="https://github.com/chriswang06/DisderoPoExcelCreator"
REPO_NAME="DisderoPoExcelCreator"

echo -e "${BLUE}=====================================${NC}"
echo -e "${BLUE}   DisderoPoExcelCreator - Quick Installer    ${NC}"
echo -e "${BLUE}=====================================${NC}"
echo ""

# Check if git is installed
if ! command -v git &> /dev/null; then
    echo -e "${RED}[!] Git is not installed${NC}"
    echo "Installing git..."

    if [[ "$OSTYPE" == "linux-gnu"* ]]; then
        if command -v apt-get &> /dev/null; then
            sudo apt-get update && sudo apt-get install -y git
        elif command -v yum &> /dev/null; then
            sudo yum install -y git
        elif command -v pacman &> /dev/null; then
            sudo pacman -S --noconfirm git
        fi
    elif [[ "$OSTYPE" == "darwin"* ]]; then
        if command -v brew &> /dev/null; then
            brew install git
        else
            echo -e "${RED}Please install Homebrew first: https://brew.sh${NC}"
            exit 1
        fi
    fi
fi

# Determine installation directory
DEFAULT_DIR="$HOME/po-processor"
echo -e "${GREEN}Where would you like to install PO Processor?${NC}"
echo -e "Default: ${YELLOW}$DEFAULT_DIR${NC}"
read -p "Press Enter for default or enter path: " INSTALL_DIR

if [ -z "$INSTALL_DIR" ]; then
    INSTALL_DIR="$DEFAULT_DIR"
fi

# Expand ~ to home directory
INSTALL_DIR="${INSTALL_DIR/#\~/$HOME}"

# Check if directory exists
if [ -d "$INSTALL_DIR" ]; then
    echo -e "${YELLOW}[!] Directory $INSTALL_DIR already exists${NC}"
    read -p "Remove existing installation and continue? (y/n): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo "Removing existing installation..."
        rm -rf "$INSTALL_DIR"
    else
        echo "Installation cancelled"
        exit 1
    fi
fi

# Clone repository
echo -e "${GREEN}Cloning repository...${NC}"
git clone "$REPO_URL.git" "$INSTALL_DIR"

# Change to installation directory
cd "$INSTALL_DIR"

# Check if install.sh exists
if [ -f "install.sh" ]; then
    echo -e "${GREEN}Running installation script...${NC}"
    chmod +x install.sh
    ./install.sh
else
    echo -e "${YELLOW}[!] install.sh not found, performing basic setup...${NC}"

    # Basic installation if install.sh is missing
    if command -v python3 &> /dev/null; then
        python3 -m venv venv
        source venv/bin/activate
        pip install --upgrade pip

        if [ -f "requirements.txt" ]; then
            pip install -r requirements.txt
        fi

        echo -e "${GREEN}Basic installation complete${NC}"
    else
        echo -e "${RED}Python 3 is required but not installed${NC}"
        exit 1
    fi
fi

# Create desktop entry for Linux
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    DESKTOP_FILE="$HOME/.local/share/applications/po-processor.desktop"
    mkdir -p "$HOME/.local/share/applications"

    cat > "$DESKTOP_FILE" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=PO Processor
Comment=Process Purchase Orders
Exec=$INSTALL_DIR/run.sh
Icon=application-pdf
Terminal=false
Categories=Office;
Path=$INSTALL_DIR
EOF

    chmod +x "$DESKTOP_FILE"
    echo -e "${GREEN}Desktop entry created${NC}"
fi

# Create shell alias
echo -e "${GREEN}Creating command alias...${NC}"
SHELL_RC="$HOME/.bashrc"
if [[ "$SHELL" == *"zsh"* ]]; then
    SHELL_RC="$HOME/.zshrc"
fi

# Add alias if it doesn't exist
if ! grep -q "alias po-processor" "$SHELL_RC" 2>/dev/null; then
    echo "" >> "$SHELL_RC"
    echo "# PO Processor alias" >> "$SHELL_RC"
    echo "alias po-processor='cd $INSTALL_DIR && ./run.sh'" >> "$SHELL_RC"
    echo -e "${GREEN}Added 'po-processor' command to $SHELL_RC${NC}"
fi

echo ""
echo -e "${GREEN}=====================================${NC}"
echo -e "${GREEN}    Installation Complete!           ${NC}"
echo -e "${GREEN}=====================================${NC}"
echo ""
echo -e "Installation directory: ${YELLOW}$INSTALL_DIR${NC}"
echo ""
echo -e "${BLUE}To run PO Processor:${NC}"
echo -e "  1. ${YELLOW}cd $INSTALL_DIR && ./run.sh${NC}"
echo -e "  2. ${YELLOW}po-processor${NC} (after restarting terminal)"
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    echo -e "  3. Launch from your applications menu"
fi
echo ""
echo -e "${BLUE}To update in the future:${NC}"
echo -e "  ${YELLOW}cd $INSTALL_DIR && git pull${NC}"
echo ""