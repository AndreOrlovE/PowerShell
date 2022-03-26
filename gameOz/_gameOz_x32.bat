@echo off
set CURDIR=%CD%
echo "Starting.."
%SystemRoot%\sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass %CURDIR%\gameOz.ps1 -cmd win -config %CURDIR%\config.xml
exit
