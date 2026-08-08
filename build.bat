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

REM Read version from VERSION file (fallback to main.py)
set "APP_VERSION=unknown"
if exist "VERSION" (
  set /p APP_VERSION=<VERSION
)

echo.
echo Building media-monitor v%APP_VERSION% ...
set "EXE_NAME=media-monitor-v%APP_VERSION%"
REM onedir: starts in seconds (onefile unpacks to %%TEMP%% every launch — minutes)
pyinstaller --noconfirm --clean --onedir --windowed --name "%EXE_NAME%" ^
  --distpath dist ^
  --workpath build ^
  --contents-directory _internal ^
  --add-data "certs;certs" ^
  --hidden-import gui ^
  main.py
if errorlevel 1 exit /b 1

REM Flatten onedir app folder into dist\ so configs sit next to the exe
if exist "dist\%EXE_NAME%\" (
  robocopy "dist\%EXE_NAME%" "dist" /E /MOVE >nul
  if exist "dist\%EXE_NAME%\" rd /s /q "dist\%EXE_NAME%"
)

REM Drop unversioned copy if an older build left it behind.
if exist "dist\media-monitor.exe" del /Q "dist\media-monitor.exe" >nul 2>nul

echo.
echo Building user instruction DOCX...
python build_instruction_docx.py
if errorlevel 1 (
  echo WARNING: failed to build ИНСТРУКЦИЯ.docx
)

echo.
echo Preparing dist folder...
if not exist "dist\reports" mkdir "dist\reports"

REM Always refresh examples (safe to overwrite) — образец настроек обновляется всегда
copy /Y "sites.example.txt" "dist\sites.example.txt" >nul
copy /Y "words.example.txt" "dist\words.example.txt" >nul
copy /Y "settings.example.txt" "dist\settings.example.txt" >nul
copy /Y "exclude.example.txt" "dist\exclude.example.txt" >nul
if exist "ИНСТРУКЦИЯ.txt" copy /Y "ИНСТРУКЦИЯ.txt" "dist\ИНСТРУКЦИЯ.txt" >nul
if exist "ИНСТРУКЦИЯ.docx" copy /Y "ИНСТРУКЦИЯ.docx" "dist\ИНСТРУКЦИЯ.docx" >nul
if exist "VERSION" copy /Y "VERSION" "dist\VERSION" >nul

REM Create local configs from examples ONLY if missing — never overwrite user files
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
if not exist "dist\exclude.txt" (
  copy /Y "exclude.example.txt" "dist\exclude.txt" >nul
  echo Created dist\exclude.txt from example
) else (
  echo Kept existing dist\exclude.txt
)

REM Remove legacy singular setting.txt if present (now only settings.txt)
if exist "dist\setting.txt" (
  del /Q "dist\setting.txt" >nul 2>nul
  echo Removed legacy dist\setting.txt ^(use settings.txt^)
)

echo.
echo ============================================
echo  Build OK: media-monitor v%APP_VERSION%
echo  Output: dist\%EXE_NAME%.exe  (+ _internal\)
echo ============================================
echo   %EXE_NAME%.exe
echo   _internal\   ^(библиотеки — не удалять^)
echo   settings.example.txt  ^(образец, обновляется^)
echo   settings.txt / sites.txt / words.txt / exclude.txt  ^(ваши, не затираются^)
echo   ИНСТРУКЦИЯ.docx / ИНСТРУКЦИЯ.txt / INSTRUCTION.*
echo   reports\
echo Copy the whole dist\ folder anywhere ^(USB OK^).
echo Do NOT launch from inside the zip — unpack first.
echo.
pause
endlocal
