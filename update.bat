@echo off
REM Double-click to update from the newest *-SKED.xls in this folder,
REM or drag a .xls file onto this .bat to use that file.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1" "%~1"
