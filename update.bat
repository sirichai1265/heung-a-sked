@echo off
chcp 65001 >nul
REM Double-click  -> full update from the newest *-SKED.xls in this folder.
REM Drag a .xls   -> full update from that file.
REM For a partial (some-vessels-only) SKED, use update-partial.bat instead.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1" -Mode full -File "%~1"
