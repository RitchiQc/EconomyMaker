@echo off
REM Runs the PowerShell script at the path you provided.
REM Update the PS1_PATH variable if the script is moved.

set PS1_PATH="C:\Users\cedri\Bureau\Minecraft 2026\EconomieMaker\Analyse-Transactions.ps1"

REM Run PowerShell with ExecutionPolicy Bypass
powershell -NoProfile -ExecutionPolicy Bypass -File %PS1_PATH%

pause