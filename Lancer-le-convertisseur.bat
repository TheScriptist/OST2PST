@echo off
rem Lance le convertisseur OST vers PST (fenetre Windows)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Convertir-OST-en-PST.ps1"
if errorlevel 1 pause
