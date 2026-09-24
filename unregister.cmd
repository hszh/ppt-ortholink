@echo off
setlocal
cd /d "%~dp0"
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\OrthoLink.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CLSID\{7C3B6C2E-5A1D-4F0B-9B7E-0D2A6E4F1A11}" /f >nul 2>&1
reg delete "HKCU\Software\Classes\OrthoLink.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Classes\Record" /v "{3E9C1D4A-2B7F-4E6A-8C11-5F0D9A7B2C33}" /f >nul 2>&1
echo == unregistered. Restart PowerPoint.
