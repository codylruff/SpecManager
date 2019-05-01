@ECHO OFF
REM PowerShell.exe -ExecutionPolicy Bypass -Command "Get-Process excel –ea 0 | Where-Object { $_.MainWindowTitle –like ‘*Spec Manager*’ } | Stop-Process"
PowerShell.exe -ExecutionPolicy Bypass -Command "& '%~dpn0.ps1'"
PAUSE