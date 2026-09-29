@echo off
rem  Adobe Campaign Package Splitter - launcher (Windows PowerShell 5.1 is built into Windows 10/11 and Server 2016+)
rem  Double-click            : opens the window
rem  Drag package(s) onto me : opens the window with those files filled in
rem  Command line            : Split-ACPackage.bat -InputFile "C:\path\package.xml" -MaxSizeMB 10
rem                            Split-ACPackage.bat -Analyze -InputFile "C:\path\a.xml" "C:\path\b.xml"
setlocal
set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0ACPackageSplitter.ps1"
if not exist "%SCRIPT%" (
  echo ACPackageSplitter.ps1 must be in the same folder as this .bat file.
  pause
  exit /b 1
)
set "FIRST=%~1"
if not defined FIRST goto gui
if "%FIRST:~0,1%"=="-" goto cli

:gui
start "" "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%SCRIPT%" -Gui %*
exit /b 0

:cli
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%" %*
exit /b %ERRORLEVEL%
