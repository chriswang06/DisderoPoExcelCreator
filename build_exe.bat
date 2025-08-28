@echo off
echo Building DisderoPoExcelCreator executable...

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist DisderoPoExcelCreator.spec del DisderoPoExcelCreator.spec

REM Create the executable with PyInstaller
pyinstaller ^
    --name="DisderoPoExcelCreator" ^
    --onefile ^
    --windowed ^
    --add-data="productslist.xlsx;." ^
    --add-data="src;src" ^
    --hidden-import="pytesseract" ^
    --hidden-import="PIL" ^
    --hidden-import="pdf2image" ^
    --hidden-import="openpyxl" ^
    --hidden-import="pandas" ^
    --hidden-import="tkinter" ^
    --clean ^
    gui_app.py

echo.
echo Build complete! Check the dist\ folder for DisderoPoExcelCreator.exe
pause