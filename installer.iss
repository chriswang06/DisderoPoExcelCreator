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

[Messages]
SetupWindowTitle=DisderoPoExcelCreator Setup
WelcomeLabel2=This will install DisderoPoExcelCreator and all required dependencies (Python, Tesseract, Poppler).

[Files]
; Your application files
Source: "src\*"; DestDir: "{app}\src"; Flags: recursesubdirs
Source: "gui_app.py"; DestDir: "{app}"
Source: "main.py"; DestDir: "{app}"
Source: "productslist.xlsx"; DestDir: "{app}"
Source: "requirements.txt"; DestDir: "{app}"

; Embedded installers
Source: "deps\python-3.11.7-amd64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "deps\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "deps\Release-25.07.0-0.zip"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
; Install Python
Filename: "{tmp}\python-3.11.7-amd64.exe"; Parameters: "/quiet InstallAllUsers=1 PrependPath=1"; StatusMsg: "Installing Python..."; Flags: waituntilterminated

; Install Tesseract
Filename: "{tmp}\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; Parameters: "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; StatusMsg: "Installing Tesseract OCR..."; Flags: waituntilterminated

; Filename: "{tmp}\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; Parameters: "/S"; StatusMsg: "Installing Tesseract OCR..."; Flags: waituntilterminated

; Extract Poppler
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""Expand-Archive -Path '{tmp}\Release-25.07.0-0.zip' -DestinationPath '{commonpf}\' -Force"""; StatusMsg: "Installing Poppler..."; Flags: waituntilterminated runhidden

; Setup Python environment
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && python -m venv venv && venv\Scripts\pip.exe install --upgrade pip && venv\Scripts\pip.exe install -r requirements.txt"""; StatusMsg: "Setting up Python environment..."; Flags: waituntilterminated runhidden

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  RunBatPath: string;
  RunBatContent: TArrayOfString;
begin
  if CurStep = ssPostInstall then
  begin
    // Create run.bat after installation
    RunBatPath := ExpandConstant('{app}\run.bat');
    SetArrayLength(RunBatContent, 4);
    RunBatContent[0] := '@echo off';
    RunBatContent[1] := 'cd /d "%~dp0"';
    RunBatContent[2] := 'call venv\Scripts\activate.bat';
    RunBatContent[3] := 'python gui_app.py';
    SaveStringsToFile(RunBatPath, RunBatContent, False);
  end;
end;

[Registry]
; Add Poppler to PATH
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{commonpf}\poppler-25.07.0\Library\bin"; Flags: preservestringtype

[UninstallDelete]
Type: filesandordirs; Name: "{app}\venv"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: files; Name: "{app}\run.bat"