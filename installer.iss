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
; Install Python with explicit install directory
Filename: "{tmp}\python-3.11.7-amd64.exe"; Parameters: "/quiet InstallAllUsers=1 PrependPath=1 TargetDir=""C:\Program Files\Python311"""; StatusMsg: "Installing Python 3.11.7..."; Flags: waituntilterminated

; Wait and verify Python installation
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""Start-Sleep -Seconds 5; if (Test-Path 'C:\Program Files\Python311\python.exe') {{ exit 0 }} else {{ exit 1 }}"""; StatusMsg: "Verifying Python installation..."; Flags: waituntilterminated

; Install Tesseract
Filename: "{tmp}\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; Parameters: "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; StatusMsg: "Installing Tesseract OCR..."; Flags: waituntilterminated

; Extract Poppler
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""Expand-Archive -Path '{tmp}\Release-25.07.0-0.zip' -DestinationPath '{commonpf}\' -Force"""; StatusMsg: "Installing Poppler..."; Flags: waituntilterminated runhidden

; Create virtual environment with explicit Python path and error checking
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""cd '{app}'; & 'C:\Program Files\Python311\python.exe' -m venv venv 2>&1 | Out-File '{tmp}\venv_create.log'; if (Test-Path '{app}\venv\Scripts\python.exe') {{ exit 0 }} else {{ Get-Content '{tmp}\venv_create.log'; exit 1 }}"""; StatusMsg: "Creating Python virtual environment..."; Flags: waituntilterminated

; Upgrade pip in the virtual environment
Filename: "{app}\venv\Scripts\python.exe"; Parameters: "-m pip install --upgrade pip"; StatusMsg: "Upgrading pip..."; Flags: waituntilterminated

; Install requirements with full path
Filename: "{app}\venv\Scripts\pip.exe"; Parameters: "install -r ""{app}\requirements.txt"""; StatusMsg: "Installing Python packages (this may take several minutes)..."; Flags: waituntilterminated

; Verify installation
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""& '{app}\venv\Scripts\python.exe' -c 'import pytesseract, PIL, pdf2image, openpyxl, pandas; print(\"All packages installed successfully\")' 2>&1"""; StatusMsg: "Verifying package installation..."; Flags: waituntilterminated

[Code]
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  // Check if Python is already installed and remove it if needed
  if RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Python\PythonCore\3.11') then
  begin
    MsgBox('Python 3.11 is already installed. The installer will use the existing installation.', mbInformation, MB_OK);
  end;
  Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  RunBatPath: string;
  RunBatContent: TArrayOfString;
  VenvPath: string;
  ErrorMsg: string;
begin
  if CurStep = ssPostInstall then
  begin
    // Verify venv was created
    VenvPath := ExpandConstant('{app}\venv');
    if not DirExists(VenvPath) then
    begin
      ErrorMsg := 'Virtual environment was not created. Please check Python installation.';
      MsgBox(ErrorMsg, mbError, MB_OK);
      // Try to create it one more time using cmd
      Exec('cmd.exe', '/c cd /d "' + ExpandConstant('{app}') + '" && python -m venv venv', '', SW_SHOW, ewWaitUntilTerminated, ResultCode);
    end;

    // Create run.bat after installation
    RunBatPath := ExpandConstant('{app}\run.bat');
    SetArrayLength(RunBatContent, 5);
    RunBatContent[0] := '@echo off';
    RunBatContent[1] := 'cd /d "%~dp0"';
    RunBatContent[2] := 'if exist venv\Scripts\activate.bat (';
    RunBatContent[3] := '    call venv\Scripts\activate.bat && python gui_app.py';
    RunBatContent[4] := ') else ( echo Virtual environment not found. Please reinstall. && pause )';
    SaveStringsToFile(RunBatPath, RunBatContent, False);
  end;
end;

[Registry]
; Add Python to PATH
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};C:\Program Files\Python311;C:\Program Files\Python311\Scripts"; Flags: preservestringtype

; Add Poppler to PATH
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{commonpf}\poppler-25.07.0\Library\bin"; Flags: preservestringtype

[UninstallDelete]
Type: filesandordirs; Name: "{app}\venv"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: files; Name: "{app}\run.bat"