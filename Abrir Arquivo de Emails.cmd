@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\Start-OutlookArchive.ps1"
if errorlevel 1 pause
