@echo off
title Gold Procurement Web UI
cd /d "%~dp0"

echo =====================================================
echo   Gold Procurement Report Converter - Web UI
echo =====================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.x
    pause & exit /b 1
)

echo  Adding Windows Firewall rule for port 5000 (requires Admin)...
netsh advfirewall firewall add rule name="Gold Procurement App" dir=in action=allow protocol=TCP localport=5000 >nul 2>&1

echo.
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /C:"IPv4 Address"') do (
    set LAN_IP=%%A
    goto :found
)
:found
set LAN_IP=%LAN_IP: =%

echo  Server starting...
echo.
echo  ┌─────────────────────────────────────────────────┐
echo  │  LOCAL (this PC):  http://localhost:5000        │
echo  │  OFFICE NETWORK:   http://%LAN_IP%:5000   │
echo  └─────────────────────────────────────────────────┘
echo.
echo  Share the OFFICE NETWORK link with your colleagues.
echo  Press Ctrl+C to stop the server.
echo =====================================================
echo.

python app.py

pause
