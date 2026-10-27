@echo off
setlocal
cd /d "%~dp0.."
set FX=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set GAC=%WINDIR%\assembly
for /f "usebackq delims=" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -find MSBuild\**\Roslyn\csc.exe`) do set CSC=%%i
if not defined CSC ( echo Install Visual Studio Build Tools first. & exit /b 1 )
if not exist build\OrthoLink.AddIn.dll ( echo Run build.cmd first. & exit /b 1 )
"%CSC%" -nologo -out:build\OrthoLink.PolygonTests.exe ^
  -r:build\OrthoLink.AddIn.dll -r:"%FX%\System.Drawing.dll" ^
  -link:"%GAC%\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll" ^
  -link:"%GAC%\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll" ^
  -link:"%GAC%\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll" ^
  src\OrthoLink.AddIn\PolygonGeometry.cs tests\PolygonTests.cs || exit /b 1
build\OrthoLink.PolygonTests.exe %*
exit /b %ERRORLEVEL%
