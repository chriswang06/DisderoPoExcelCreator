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
; Install Python - will skip if already installed via Check function
Filename: "{tmp}\python-3.11.7-amd64.exe"; Parameters: "/quiet InstallAllUsers=1 PrependPath=1 TargetDir=""C:\Program Files\Python311"""; StatusMsg: "Installing Python 3.11.7..."; Flags: waituntilterminated; Check: ShouldInstallPython

; Wait for Python installation to complete
Filename: "{cmd}"; Parameters: "/c ping 127.0.0.1 -n 3 > nul"; StatusMsg: "Waiting for installation to complete..."; Flags: waituntilterminated runhidden

; Install Tesseract - using /S for NSIS installer
Filename: "{tmp}\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; Parameters: "/S"; StatusMsg: "Installing Tesseract OCR..."; Flags: waituntilterminated

; Extract Poppler
Filename: "powershell.exe"; Parameters: "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command ""Expand-Archive -Path '{tmp}\Release-25.07.0-0.zip' -DestinationPath '{commonpf}' -Force"""; StatusMsg: "Installing Poppler..."; Flags: waituntilterminated runhidden

; Create virtual environment with the best available Python
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && ({code:GetPythonPath} -m venv venv || python -m venv venv)"""; StatusMsg: "Creating virtual environment..."; Flags: waituntilterminated runhidden

; Upgrade pip
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && venv\Scripts\python.exe -m pip install --upgrade pip --quiet --no-warn-script-location"""; StatusMsg: "Upgrading pip..."; Flags: waituntilterminated runhidden

; Install requirements
Filename: "{cmd}"; Parameters: "/c ""cd /d ""{app}"" && venv\Scripts\pip.exe install --no-cache-dir --quiet -r requirements.txt"""; StatusMsg: "Installing Python packages (this may take a few minutes)..."; Flags: waituntilterminated runhidden

[Code]
function ShouldInstallPython(): Boolean;
var
  Version: String;
  ExitCode: Integer;
begin
  // Check if Python 3.11 is already installed
  // First check the exact path we would install to
  if FileExists('C:\Program Files\Python311\python.exe') then
  begin
    Log('Python 3.11 found at expected installation path');
    Result := False;
    Exit;
  end;

  // Check if any Python 3.11 is available via PATH
  if Exec('cmd.exe', '/c python --version 2>&1 | findstr "3.11"', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) and (ExitCode = 0) then
  begin
    Log('Python 3.11 found in PATH');
    Result := False;
    Exit;
  end;

  // Check registry for Python 3.11
  if RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Python\PythonCore\3.11') then
  begin
    Log('Python 3.11 found in registry');
    Result := False;
    Exit;
  end;

  // No Python 3.11 found, need to install
  Log('Python 3.11 not found, will install');
  Result := True;
end;

function GetPythonPath(Param: String): String;
begin
  // Return the path to Python executable
  if FileExists('C:\Program Files\Python311\python.exe') then
    Result := '"C:\Program Files\Python311\python.exe"'
  else
    Result := 'python';  // Fallback to system Python
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  RunBatPath: string;
  RunBatContent: TArrayOfString;
  VenvPath: string;
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    // Verify venv was created
    VenvPath := ExpandConstant('{app}\venv');
    if DirExists(VenvPath) then
    begin
      Log('Virtual environment created successfully at: ' + VenvPath);
    end
    else
    begin
      Log('WARNING: Virtual environment not found, attempting alternative creation methods');

      // Try different Python commands
      if not Exec('py', '-3.11 -m venv "' + VenvPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
        if not Exec('py', '-3 -m venv "' + VenvPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
          Exec('python', '-m venv "' + VenvPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    end;

    // Create run.bat
    RunBatPath := ExpandConstant('{app}\run.bat');
    SetArrayLength(RunBatContent, 10);
    RunBatContent[0] := '@echo off';
    RunBatContent[1] := 'cd /d "%~dp0"';
    RunBatContent[2] := '';
    RunBatContent[3] := 'if not exist venv\Scripts\activate.bat (';
    RunBatContent[4] := '    echo ERROR: Virtual environment not found.';
    RunBatContent[5] := '    echo Please reinstall the application.';
    RunBatContent[6] := '    pause';
    RunBatContent[7] := '    exit /b 1';
    RunBatContent[8] := ')';
    RunBatContent[9] := '';

    SetArrayLength(RunBatContent, 12);
    RunBatContent[10] := 'call venv\Scripts\activate.bat';
    RunBatContent[11] := 'python gui_app.py';

    SaveStringsToFile(RunBatPath, RunBatContent, False);
  end;
end;

[Registry]
; Add Python to PATH only if we installed it
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};C:\Program Files\Python311;C:\Program Files\Python311\Scripts"; Flags: preservestringtype; Check: ShouldInstallPython

; Add Poppler to PATH
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{commonpf}\poppler-25.07.0\Library\bin"; Flags: preservestringtype

[UninstallDelete]
Type: filesandordirs; Name: "{app}\venv"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: files; Name: "{app}\run.bat"