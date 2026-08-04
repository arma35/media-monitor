@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  python -m venv .venv
  if errorlevel 1 exit /b 1
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 exit /b 1

echo.
echo Building portable exe...
pyinstaller --noconfirm --clean --onefile --name media-monitor ^
  --distpath dist ^
  --workpath build ^
  main.py
if errorlevel 1 exit /b 1

echo.
echo Preparing dist folder...
if not exist "dist\reports" mkdir "dist\reports"

REM Always refresh examples (safe to overwrite)
copy /Y "sites.example.txt" "dist\sites.example.txt" >nul
copy /Y "words.example.txt" "dist\words.example.txt" >nul
copy /Y "settings.example.txt" "dist\settings.example.txt" >nul

REM Create local configs from examples ONLY if missing — never overwrite
if not exist "dist\sites.txt" (
  copy /Y "sites.example.txt" "dist\sites.txt" >nul
  echo Created dist\sites.txt from example
) else (
  echo Kept existing dist\sites.txt
)
if not exist "dist\words.txt" (
  copy /Y "words.example.txt" "dist\words.txt" >nul
  echo Created dist\words.txt from example
) else (
  echo Kept existing dist\words.txt
)
if not exist "dist\settings.txt" (
  copy /Y "settings.example.txt" "dist\settings.txt" >nul
  echo Created dist\settings.txt from example
) else (
  echo Kept existing dist\settings.txt
)

echo.
echo Done. Portable folder: dist\
echo   media-monitor.exe
echo   *.example.txt                          ^(samples in git^)
echo   sites.txt / words.txt / settings.txt   ^(local, not overwritten^)
echo   reports\
echo Copy the whole dist\ folder anywhere ^(USB OK^).
endlocal
