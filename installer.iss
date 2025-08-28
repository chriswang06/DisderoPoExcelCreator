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

; Wait for Python installation to complete
Filename: "{cmd}"; Parameters: "/c timeout /t 5"; StatusMsg: "Waiting for Python installation to complete..."; Flags: waituntilterminated runhidden

; Install Tesseract
Filename: "{tmp}\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; Parameters: "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; StatusMsg: "Installing Tesseract OCR..."; Flags: waituntilterminated

; Extract Poppler
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -Command ""Expand-Archive -Path '{tmp}\Release-25.07.0-0.zip' -DestinationPath '{commonpf}' -Force"""; StatusMsg: "Installing Poppler..."; Flags: waituntilterminated runhidden

; Create virtual environment using cmd instead of PowerShell (simpler escaping)
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && ""C:\Program Files\Python311\python.exe"" -m venv venv"""; StatusMsg: "Creating Python virtual environment..."; Flags: waituntilterminated

; Verify venv creation
Filename: "{cmd}"; Parameters: "/c ""if exist ""{app}\venv\Scripts\python.exe"" (echo Venv created successfully) else (echo ERROR: Venv creation failed && exit 1)"""; StatusMsg: "Verifying virtual environment..."; Flags: waituntilterminated

; Upgrade pip in the virtual environment
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && venv\Scripts\python.exe -m pip install --upgrade pip"""; StatusMsg: "Upgrading pip..."; Flags: waituntilterminated

; Install requirements
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && venv\Scripts\pip.exe install -r requirements.txt"""; StatusMsg: "Installing Python packages (this may take several minutes)..."; Flags: waituntilterminated

; Verify package installation
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && venv\Scripts\python.exe -c ""import pytesseract, PIL, pdf2image, openpyxl, pandas; print('All packages installed successfully')"""""; StatusMsg: "Verifying package installation..."; Flags: waituntilterminated

[Code]
var
  ResultCode: Integer;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  // Check if Python is already installed
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
      ErrorMsg := 'Virtual environment was not created. Attempting to create it again...';
      MsgBox(ErrorMsg, mbInformation, MB_OK);
      // Try to create it one more time
      Exec('cmd.exe', '/c cd /d "' + ExpandConstant('{app}') + '" && python -m venv venv', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

      // Check again
      if not DirExists(VenvPath) then
      begin
        MsgBox('Failed to create virtual environment. Please ensure Python is installed correctly.', mbError, MB_OK);
      end;
    end;

    // Create run.bat after installation
    RunBatPath := ExpandConstant('{app}\run.bat');
    SetArrayLength(RunBatContent, 9);
    RunBatContent[0] := '@echo off';
    RunBatContent[1] := 'cd /d "%~dp0"';
    RunBatContent[2] := '';
    RunBatContent[3] := 'if not exist venv\Scripts\activate.bat (';
    RunBatContent[4] := '    echo ERROR: Virtual environment not found.';
    RunBatContent[5] := '    echo Please reinstall the application.';
    RunBatContent[6] := '    pause';
    RunBatContent[7] := '    exit /b 1';
    RunBatContent[8] := ')';

    // Add the activation and run commands
    SetArrayLength(RunBatContent, 12);
    RunBatContent[9] := '';
    RunBatContent[10] := 'call venv\Scripts\activate.bat';
    RunBatContent[11] := 'python gui_app.py';

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