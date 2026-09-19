@echo off
rem Starts the BMS HRMS biometric sync agent and restarts it automatically if it stops.
cd /d "%~dp0"
:loop
node sync-agent.js
echo Agent stopped. Restarting in 30 seconds...
timeout /t 30 /nobreak >nul
goto loop
