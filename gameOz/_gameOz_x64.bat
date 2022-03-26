@echo off
set CURDIR=%CD%
echo "Starting.."
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass %CURDIR%\gameOz.ps1 -cmd win -config %CURDIR%\config.xml
exit
