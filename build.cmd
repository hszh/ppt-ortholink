@echo off
setlocal
cd /d "%~dp0"

set FX=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set GAC=%WINDIR%\assembly
set PIA=%GAC%\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll
set OFFICE=%GAC%\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll
set EXT=%GAC%\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll
set STDOLE=%GAC%\GAC\stdole\7.0.3300.0__b03f5f7f11d50a3a\stdole.dll
set VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe

for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -find MSBuild\**\Roslyn\csc.exe`) do set CSC=%%i
if not defined CSC (
  echo Could not find csc.exe. Install Visual Studio Build Tools with the MSBuild workload.
  exit /b 1
)
for /f tokens^=2^ delims^=^" %%v in ('findstr /c:"Display = " src\Version.cs') do set VER=%%v
if not defined VER ( echo Could not read the version from src\Version.cs & exit /b 1 )

echo == template + icons
python tools\build_template.py || exit /b 1
rem the icons in assets\ are committed; regenerate them only where the local script exists
if exist tools\build_icon.py ( python tools\build_icon.py || exit /b 1 )

echo == compile add-in %VER%
if not exist build mkdir build
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | "%WINDIR%\System32\find.exe" /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
  echo PowerPoint is running and holds the add-in file. Close PowerPoint, then run build.cmd again.
  exit /b 1
)
"%CSC%" -nologo -target:library -platform:anycpu -optimize+ -debug:pdbonly -unsafe ^
  -out:build\OrthoLink.AddIn.dll ^
  -r:"%FX%\System.Windows.Forms.dll" -r:"%FX%\System.Drawing.dll" ^
  -link:"%PIA%" -link:"%OFFICE%" -link:"%EXT%" -link:"%STDOLE%" ^
  -resource:src\OrthoLink.AddIn\Ribbon.xml,OrthoLink.Ribbon.xml ^
  -resource:assets\icon-32.png,OrthoLink.icon32.png ^
  -resource:assets\icon-16.png,OrthoLink.icon16.png ^
  src\Version.cs src\OrthoLink.AddIn\*.cs || exit /b 1
copy /y src\OrthoLink.AddIn\template.pptx build\ >nul

echo == compile setup program
if not exist dist mkdir dist
"%CSC%" -nologo -target:winexe -platform:anycpu -optimize+ ^
  -win32icon:assets\icon.ico ^
  -out:dist\OrthoLink-Setup-%VER%.exe ^
  -r:"%FX%\System.Windows.Forms.dll" -r:"%FX%\System.Drawing.dll" ^
  -resource:build\OrthoLink.AddIn.dll,OrthoLink.Setup.OrthoLink.AddIn.dll ^
  -resource:build\template.pptx,OrthoLink.Setup.template.pptx ^
  src\Version.cs src\OrthoLink.Setup\*.cs || exit /b 1

echo == done
echo    add-in : build\OrthoLink.AddIn.dll  (register.cmd registers this copy for development)
echo    setup  : dist\OrthoLink-Setup-%VER%.exe  (give this file to users)
