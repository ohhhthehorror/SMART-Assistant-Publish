@echo off
REM ============================================
REM SMART Assistant EXE Builder
REM Automated build script for Windows
REM ============================================

echo.
echo ============================================
echo SMART Assistant v0.2.1 - EXE Builder
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.12+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Python found
echo.

REM Check if required files exist
if not exist "SMART_ASSISTANT_FINAL.py" (
    echo [ERROR] SMART_ASSISTANT_FINAL.py not found!
    pause
    exit /b 1
)

echo [OK] Main script found
echo.

REM Install dependencies
echo ============================================
echo Step 1: Installing Dependencies
echo ============================================
echo.

pip install pyinstaller requests Pillow --quiet --upgrade
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)

echo [OK] Dependencies installed
echo.

REM Create ICO file if logo exists
if exist "smartlogosmall.png" (
    echo ============================================
    echo Step 2: Creating Icon File
    echo ============================================
    echo.
    
    python -c "from PIL import Image; img = Image.open('smartlogosmall.png'); img.save('smartlogo.ico', format='ICO', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)]); print('[OK] Icon created')"
    
    echo.
) else (
    echo [WARNING] smartlogosmall.png not found - building without icon
    echo.
)

REM Clean previous builds
echo ============================================
echo Step 3: Cleaning Previous Builds
echo ============================================
echo.

if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

echo [OK] Cleanup complete
echo.

REM Build the EXE
echo ============================================
echo Step 4: Building EXE
echo ============================================
echo.
echo This may take 2-5 minutes...
echo.

REM Determine if we have an icon
set ICON_ARG=
if exist "smartlogo.ico" set ICON_ARG=--icon=smartlogo.ico

REM Build command - ONEFILE mode (single executable)
pyinstaller --onefile --windowed %ICON_ARG% --add-data "auto_updater.py;." --hidden-import=auto_updater --hidden-import=requests --hidden-import=selenium --hidden-import=openpyxl --hidden-import=PIL --name SMART_Assistant SMART_ASSISTANT_FINAL.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo ============================================
echo BUILD COMPLETE!
echo ============================================
echo.

if exist "dist\SMART_Assistant.exe" (
    for %%A in ("dist\SMART_Assistant.exe") do (
        set size=%%~zA
        set /a sizeMB=%%~zA / 1048576
    )
    
    echo [SUCCESS] Single EXE created successfully!
    echo.
    echo Location: dist\SMART_Assistant.exe
    echo Size: %sizeMB% MB
    echo.
    echo ============================================
    echo Next Steps:
    echo ============================================
    echo 1. Test the EXE: dist\SMART_Assistant.exe
    echo 2. Upload this single EXE to GitHub release
    echo 3. Update version.json with release info
    echo.
    
    REM Ask if user wants to open the dist folder
    set /p OPEN="Open dist folder? (Y/N): "
    if /i "%OPEN%"=="Y" explorer dist
    
) else (
    echo [ERROR] EXE not found in dist folder
    echo Build may have failed - check output above
)

echo.
pause
