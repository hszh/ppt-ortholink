@echo off
setlocal
cd /d "%~dp0"
set REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe
set DLL=%CD%\build\OrthoLink.AddIn.dll
set ADDINKEY=HKCU\Software\Microsoft\Office\PowerPoint\Addins\OrthoLink.AddIn

if not exist "%DLL%" ( echo build first: build.cmd & exit /b 1 )

echo == generate COM registration (per-user, no admin needed)
"%REGASM%" "%DLL%" /codebase /regfile:build\ortholink.reg >nul || exit /b 1
python -c "import io;p='build/ortholink.reg';t=io.open(p,encoding='utf-8').read().replace('HKEY_CLASSES_ROOT','HKEY_CURRENT_USER\\Software\\Classes');io.open('build/ortholink_user.reg','w',encoding='utf-8').write(t)" || exit /b 1
reg import build\ortholink_user.reg >nul 2>&1 || ( echo reg import failed & exit /b 1 )

echo == tell PowerPoint to load it
reg add "%ADDINKEY%" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
reg add "%ADDINKEY%" /v FriendlyName /t REG_SZ /d "OrthoLink" /f >nul
reg add "%ADDINKEY%" /v Description /t REG_SZ /d "Rounded orthogonal connectors" /f >nul
echo == registered. Restart PowerPoint.
