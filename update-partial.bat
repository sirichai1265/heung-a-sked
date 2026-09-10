@echo off
chcp 65001 >nul
REM Drag a partial SKED .xls (some vessels only, each with its full
REM rotation) onto this file to merge it into the current dataset.
if "%~1"=="" (
  echo.
  echo ลากไฟล์ SKED ^(บางเรือ^) มาวางบนไฟล์ update-partial.bat นี้
  echo.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1" -Mode merge -File "%~1"
