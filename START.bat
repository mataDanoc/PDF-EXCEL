@echo off
chcp 65001 > nul
title PDF to Excel - Server

echo.
echo  ============================================================
echo    PDF to Excel Converter - LOCAL SERVER
echo    Port: 5050  ^|  http://localhost:5050
echo  ============================================================
echo.

:: Gjej Python
where python > nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [GABIM] Python nuk u gjet! Instalo Python 3.10+
    pause
    exit /b 1
)

:: Shko te direktoria e projektit
cd /d "%~dp0"

:: Kontrollo nese porta 5050 eshte e lire
netstat -an | find ":5050 " | find "LISTENING" > nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [KUJDES] Porta 5050 eshte e zene. Duke u mbylle procesi ekzistues...
    for /f "tokens=5" %%a in ('netstat -aon ^| find ":5050 " ^| find "LISTENING"') do (
        taskkill /F /PID %%a > nul 2>&1
    )
    timeout /t 2 /nobreak > nul
)

:: Kontrollo nese porta 5051 eshte e lire (backup port)
netstat -an | find ":5051 " | find "LISTENING" > nul 2>&1
if %ERRORLEVEL% EQU 0 (
    for /f "tokens=5" %%a in ('netstat -aon ^| find ":5051 " ^| find "LISTENING"') do (
        taskkill /F /PID %%a > nul 2>&1
    )
    timeout /t 1 /nobreak > nul
)

:: Krijo dosjet nese nuk ekzistojne
if not exist "input"  mkdir input
if not exist "output" mkdir output

:: Shfaq instruksionet
echo  [OK] Duke nisur serverin...
echo.
echo  INSTRUKSIONE:
echo    1. Hap browser-in te: http://localhost:5050
echo    2. Terhiq PDF-te ne zone te ngarkimit
echo    3. Kliko "Konverto tani"
echo    4. Shkarko Excel-in e gjeneruar
echo.
echo  Shtype CTRL+C per te ndalur serverin
echo  Ose ekzekuto STOP.bat ne dritare tjeter
echo  ============================================================
echo.

:: Hap browser-in pas 2 sekondash (ne background)
start /b cmd /c "timeout /t 2 /nobreak > nul && start http://localhost:5050"

:: Starto FastAPI me uvicorn - PORT 5050
python -m uvicorn webapp.app:app --host 127.0.0.1 --port 5050 --reload --log-level info

echo.
echo  [INFO] Serveri u ndal.
pause
