@echo off
rem ===========================================================================
rem  Build TapeProcessFormGenerator.exe
rem ---------------------------------------------------------------------------
rem  Double-click this file. It packs TapeProcessFormGenerator.pyw, its
rem  libraries and logo.ico into ONE file:  dist\TapeProcessFormGenerator.exe
rem  The .exe runs on other 64-bit Windows 10/11 PCs without Python installed.
rem
rem  Only the PC that runs this build needs Python 3 (python.org installer,
rem  which provides the "py" launcher, or "python" on PATH) and internet
rem  access the first time. The build tools go into build\venv, so your own
rem  Python installation is left untouched.
rem
rem  Hand out the .exe together with IBAD-PF-009-Rev00_Tape_Process_Form.pdf
rem  in the same folder; the program picks the template up from there.
rem ===========================================================================
setlocal
rem pushd (not cd) also works when this file sits on a network share
pushd "%~dp0" || goto :failed

set "APP=TapeProcessFormGenerator"
set "ICON=logo.ico"
set "TEMPLATE=IBAD-PF-009-Rev00_Tape_Process_Form.pdf"
set "VENV=build\venv"

if not exist "%APP%.pyw" (
    echo Cannot find %APP%.pyw next to this batch file.
    goto :failed
)
if not exist "%ICON%" (
    echo Cannot find %ICON% next to this batch file.
    goto :failed
)

rem Failures are checked as "errorlevel is not 0"; the shorter "if errorlevel 1"
rem means 1 or higher and would miss negative exit codes (crashes, Python
rem install manager errors).

rem -- 1. Private Python environment with the build tools (first run only) ---
if exist "%VENV%\Scripts\python.exe" goto :have_venv

set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)
if not defined PY goto :no_python
%PY% -c "import sys" >nul 2>nul
if %errorlevel% neq 0 goto :no_python

echo Creating the build environment in %VENV% ...
%PY% -m venv "%VENV%"
if %errorlevel% neq 0 goto :failed

:have_venv
set "VPY=%VENV%\Scripts\python.exe"

echo Installing / updating PyInstaller and the program's libraries ...
"%VPY%" -m pip install --upgrade --disable-pip-version-check pip pyinstaller pandas openpyxl pypdf reportlab
if %errorlevel% neq 0 goto :failed

rem -- 2. Build the .exe ------------------------------------------------------
rem  --windowed       no black console window behind the program
rem  --icon           logo shown for the .exe in Explorer and on the taskbar
rem  --add-data       logo packed inside the .exe for the program's window
rem  --hidden-import  pandas loads openpyxl on demand; named so it is always packed
echo.
echo Building %APP%.exe (this takes a minute or two) ...
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "%APP%" ^
    --icon "%ICON%" ^
    --add-data "%ICON%;." ^
    --hidden-import openpyxl ^
    "%APP%.pyw"
if %errorlevel% neq 0 goto :failed

rem -- 3. Put the form template next to the .exe, if it is here ---------------
if exist "%TEMPLATE%" copy /y "%TEMPLATE%" "dist\%TEMPLATE%" >nul

echo.
echo ===========================================================================
echo  Done:  "%CD%\dist\%APP%.exe"
if exist "dist\%TEMPLATE%" (
    echo  The template PDF was copied next to it.
) else (
    echo  Put %TEMPLATE% in the same folder as the .exe.
)
echo ===========================================================================
pause
exit /b 0

:no_python
echo.
echo Python 3 was not found. Install it from https://www.python.org/downloads/
echo (the default installer options include the "py" launcher), then run this
echo file again.
pause
exit /b 1

:failed
echo.
echo *** Build failed - see the messages above. ***
echo If it keeps failing, delete the "build" folder and run this file again.
pause
exit /b 1
