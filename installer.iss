[Setup]
AppName=DisderoPoExcelCreator
AppVersion=1.0
AppPublisher=Your Name
DefaultDirName={autopf}\DisderoPoExcelCreator
DefaultGroupName=DisderoPoExcelCreator
OutputDir=output
OutputBaseFilename=DisderoPoExcelCreator-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
; Add uninstall info
UninstallDisplayName=DisderoPoExcelCreator
UninstallDisplayIcon={app}\DisderoPoExcelCreator.exe

[Messages]
SetupWindowTitle=DisderoPoExcelCreator Setup
WelcomeLabel2=This will install DisderoPoExcelCreator and required OCR dependencies.

[Files]
; The compiled exe from PyInstaller (from dist folder)
Source: "dist\DisderoPoExcelCreator.exe"; DestDir: "{app}"; Flags: ignoreversion

; Dependencies that still need to be installed
Source: "deps\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "deps\Release-25.07.0-0.zip"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
; Install Tesseract
Filename: "{tmp}\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; Parameters: "/S"; StatusMsg: "Installing Tesseract OCR..."; Flags: waituntilterminated

; Extract Poppler
Filename: "powershell.exe"; Parameters: "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command ""Expand-Archive -Path '{tmp}\Release-25.07.0-0.zip' -DestinationPath '{commonpf}' -Force"""; StatusMsg: "Installing Poppler..."; Flags: waituntilterminated runhidden

[Icons]
; Desktop shortcut
Name: "{commondesktop}\DisderoPoExcelCreator"; Filename: "{app}\DisderoPoExcelCreator.exe"; Tasks: desktopicon

; Start Menu shortcuts
Name: "{group}\DisderoPoExcelCreator"; Filename: "{app}\DisderoPoExcelCreator.exe"
Name: "{group}\Uninstall DisderoPoExcelCreator"; Filename: "{uninstallexe}"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"