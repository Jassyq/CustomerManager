@echo off
title Customer Price Manager v4.0 - Memory + Export Edition
color 0A

echo ════════════════════════════════════════════════════════
echo   Customer Price Manager v4.0
echo   Memory-Based + Clean Export Edition
echo ════════════════════════════════════════════════════════
echo.

REM ============================================================
REM STEP 1: Test Python Installation
REM ============================================================
echo [Step 1/4] Testing Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERROR: Python not found!
    echo.
    echo Please install Python 3.8 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

REM Display Python version
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo ✓ Python %PYTHON_VERSION% detected
echo.

REM ============================================================
REM STEP 2: Check pip Installation
REM ============================================================
echo [Step 2/4] Testing pip (package manager)...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERROR: pip not found!
    echo.
    echo Installing pip...
    python -m ensurepip --default-pip
    if errorlevel 1 (
        echo Failed to install pip automatically
        echo Please install pip manually
        pause
        exit /b 1
    )
)
echo ✓ pip is available
echo.

REM ============================================================
REM STEP 3: Check and Install Required Packages
REM ============================================================
echo [Step 3/4] Checking required packages...
echo.

REM Check if requirements.txt exists
if not exist "requirements.txt" (
    echo ⚠️  WARNING: requirements.txt not found
    echo Creating basic requirements.txt...
    (
        echo streamlit^>=1.28.0
        echo openpyxl^>=3.1.2
        echo pandas^>=2.0.0
    ) > requirements.txt
    echo ✓ Created requirements.txt
    echo.
)

REM Test if packages are installed
echo Testing package imports...
python -c "import streamlit" 2>nul
set STREAMLIT_OK=%errorlevel%

python -c "import openpyxl" 2>nul
set OPENPYXL_OK=%errorlevel%

python -c "import pandas" 2>nul
set PANDAS_OK=%errorlevel%

REM Calculate if any package is missing
set /a PACKAGES_MISSING=%STREAMLIT_OK% + %OPENPYXL_OK% + %PANDAS_OK%

if %PACKAGES_MISSING% gtr 0 (
    echo.
    echo ⚠️  Some packages are missing or outdated
    if %STREAMLIT_OK% neq 0 echo    - streamlit
    if %OPENPYXL_OK% neq 0 echo    - openpyxl
    if %PANDAS_OK% neq 0 echo    - pandas
    echo.
    echo Installing/updating packages from requirements.txt...
    echo This may take 1-2 minutes on first run...
    echo.
    
    echo Upgrading pip to latest version...
    python -m pip install --upgrade pip setuptools wheel
    echo.
    echo Installing packages - this may take 2-3 minutes...
    python -m pip install -r requirements.txt --no-build-isolation
    
    if errorlevel 1 (
        echo.
        echo ❌ ERROR: Package installation failed!
        echo.
        echo Possible reasons:
        echo   1. No internet connection
        echo   2. Administrator privileges required
        echo   3. Antivirus blocking installation
        echo.
        echo Try running this file as Administrator
        echo ^(Right-click → Run as administrator^)
        echo.
        pause
        exit /b 1
    )
    
    echo.
    echo ✓ All packages installed successfully!
) else (
    echo ✓ All required packages are already installed
)
echo.

REM ============================================================
REM STEP 4: Locate and Verify App File
REM ============================================================
echo [Step 4/4] Locating application file...

if exist "app.py" (
    set APP_FILE=app.py
    echo ✓ Found: app.py
) else (
    echo ❌ ERROR: app.py not found!
    echo.
    echo Current directory: %CD%
    echo.
    echo Files in this directory:
    dir /b *.py 2>nul
    if errorlevel 1 (
        echo    ^(No Python files found^)
    )
    echo.
    pause
    exit /b 1
)
echo.

REM ============================================================
REM Display Features and Instructions
REM ============================================================
echo ════════════════════════════════════════════════════════
echo   ✓ ALL CHECKS PASSED - READY TO START!
echo ════════════════════════════════════════════════════════
echo.
echo   KEY FEATURES:
echo   ✓ NO FILE PERMISSION ERRORS!
echo   ✓ All changes stored in memory
echo   ✓ Survives refresh and disconnect
echo   ✓ Edit at your own pace
echo   ✓ Export clean formatted Excel
echo   ✓ Professional output format
echo ════════════════════════════════════════════════════════
echo.
echo   WORKFLOW:
echo   1. Upload Excel file ^(read-only OK!^)
echo   2. Edit customers one by one
echo   3. Click "Save to Memory" for each customer
echo   4. Click "Export Clean Excel" when done
echo   5. Download your formatted file
echo.
echo ════════════════════════════════════════════════════════
echo.
echo   💡 TIPS:
echo   • Export frequently to save your work
echo   • Keep this window open while using the app
echo   • Browser will open automatically
echo   • Use Ctrl+C in this window to stop the app
echo.
echo ════════════════════════════════════════════════════════
echo.

echo Press any key to launch the application...
pause >nul

REM ============================================================
REM Launch Streamlit Application
REM ============================================================
cls
echo ════════════════════════════════════════════════════════
echo   CUSTOMER PRICE MANAGER IS NOW RUNNING
echo ════════════════════════════════════════════════════════
echo.
echo   Browser should open automatically...
echo   If not, open: http://localhost:8501
echo.
echo   ⚠️  KEEP THIS WINDOW OPEN!
echo   Press Ctrl+C to stop the application
echo.
echo ════════════════════════════════════════════════════════
echo.

python -m streamlit run app.py

REM ============================================================
REM Cleanup and Exit Message
REM ============================================================
echo.
echo.
echo ════════════════════════════════════════════════════════
echo   Application has stopped
echo ════════════════════════════════════════════════════════
echo.
echo   💾 Don't forget to export your work before closing!
echo.
echo   To restart: Double-click RUN_V4_APP.bat again
echo.
echo ════════════════════════════════════════════════════════
echo.
pause
