@echo off
title PBHRC Dashboard Local Server
echo Starting PBHRC Dashboard Server...
echo Please leave this black window open while using the dashboard.

:: Launch the browser after a 2-second delay to let the server start up
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:8080/index.html"

:: Run the PowerShell server script continuously
PowerShell -ExecutionPolicy Bypass -File "%~dp0server.ps1"
