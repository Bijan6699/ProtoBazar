@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo   ProtoBazar - Windows EXE Builder
echo ==========================================
echo.
where py >nul 2>nul
if errorlevel 1 (
  echo Python was not found.
  echo Install Python 3.11 or 3.12 from https://www.python.org/downloads/windows/
  echo During setup, enable "Add python.exe to PATH".
  pause
  exit /b 1
)
py -3.11 --version >nul 2>nul
if errorlevel 1 (
  py -3.12 --version >nul 2>nul
  if errorlevel 1 (
    echo Python 3.11 or 3.12 is required to BUILD the EXE.
    pause
    exit /b 1
  ) else (
    set PY=py -3.12
  )
) else (
  set PY=py -3.11
)
echo Creating build environment...
%PY% -m venv .venv
if errorlevel 1 goto fail
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto fail
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fail
echo.
echo Building one-folder Windows application...
python -m PyInstaller --noconfirm --clean --onedir --name ProtoBazar ^
  --collect-all streamlit ^
  --collect-all plotly ^
  --collect-all pandas ^
  --collect-all numpy ^
  --collect-all openpyxl ^
  --collect-all xlrd ^
  --collect-all defusedxml ^
  --add-data "dashboard.py;." ^
  --add-data "pricing_module.py;." ^
  --add-data "sales_analytics_module.py;." ^
  --add-data "logo.png;." ^
  launcher.py
if errorlevel 1 goto fail
copy /Y "RUN_PROTOBAZAR.bat" "dist\RUN_PROTOBAZAR.bat" >nul
if errorlevel 1 (
  echo Could not copy RUN_PROTOBAZAR.bat. Make sure it exists.
  goto fail
)
echo.
echo SUCCESS: dist\ProtoBazar is ready.
echo Test it on a Windows computer that does NOT have Python installed.
echo Then share the entire dist\ProtoBazar folder, not just the EXE.
pause
exit /b 0
:fail
echo.
echo BUILD FAILED. Read the error above.
pause
exit /b 1
