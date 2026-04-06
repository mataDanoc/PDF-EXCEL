@echo off
chcp 65001 > nul
title PDF to Excel - Server
cd /d "%~dp0"

echo.
echo  ============================================================
echo    PDF to Excel Converter
echo  ============================================================
echo.

:: Gjej Python
where python > nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [GABIM] Python nuk u gjet! Instalo Python 3.10+
    pause
    exit /b 1
)

:: Shto Tesseract OCR ne PATH nese ekziston
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    set "PATH=%PATH%;C:\Program Files\Tesseract-OCR"
)

:: Mbyll procese ekzistuese ne portat
for /f "tokens=5" %%a in ('netstat -aon ^| find ":5050 " ^| find "LISTENING" 2^>nul') do taskkill /F /PID %%a > nul 2>&1

:: Krijo dosjet
if not exist "input"  mkdir input
if not exist "output" mkdir output

echo  [1/2] Duke nisur serverin lokal ne port 5050...
start "PDF-Excel Server" /min cmd /k "cd /d "%~dp0" && python -m uvicorn webapp.app:app --host 127.0.0.1 --port 5050 --log-level info"

echo  [2/2] Duke pritur serverin te nisë...
timeout /t 4 /nobreak > nul

echo  [3/3] Duke hapur tunelin publik...
echo.
echo  ============================================================
echo   LOKAL  (ti):   http://localhost:5050
echo.
echo   PUBLIK (miqte): shiko dritaren e re te hapur
echo   Kerko rreshtin:  https://xxxxx.trycloudflare.com
echo   Kopjo ate link dhe dergoja miqve!
echo  ============================================================
echo.

:: Hap browser-in lokal
start http://localhost:5050

:: Starto tunelin Cloudflare ne dritare te re
start "Cloudflare Tunnel - LINK PUBLIK" cmd /k "echo. && echo  Lidhu... && echo. && cloudflared.exe tunnel --url http://localhost:5050 2>&1"

echo  Te dy sherbimet po punojne.
echo  Kur te mbarosh, ekzekuto STOP.bat
echo.
pause
