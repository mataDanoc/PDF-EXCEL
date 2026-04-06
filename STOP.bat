@echo off
chcp 65001 > nul
title PDF to Excel - STOP

echo.
echo  ============================================================
echo    PDF to Excel Converter - NDALO SERVERIN
echo  ============================================================
echo.

:: Ndalo te gjitha proceset ne port 5050
echo  Duke kerkuar proceset ne port 5050...
set FOUND=0

for /f "tokens=5" %%a in ('netstat -aon ^| find ":5050 " ^| find "LISTENING" 2^>nul') do (
    echo  [STOP] Duke mbylle PID: %%a (port 5050)
    taskkill /F /PID %%a > nul 2>&1
    set FOUND=1
)

:: Ndalo te gjitha proceset ne port 5051 (backup port)
for /f "tokens=5" %%a in ('netstat -aon ^| find ":5051 " ^| find "LISTENING" 2^>nul') do (
    echo  [STOP] Duke mbylle PID: %%a (port 5051)
    taskkill /F /PID %%a > nul 2>&1
    set FOUND=1
)

:: Ndalo cloudflared
echo  Duke ndalur tunelin Cloudflare...
taskkill /F /IM cloudflared.exe > nul 2>&1

:: Ndalo edhe proceset uvicorn/python ne emer
echo  Duke kerkuar proceset uvicorn...
for /f "tokens=2" %%a in ('tasklist /fi "imagename eq python.exe" /fo table /nh 2^>nul') do (
    wmic process where "processid=%%a" get commandline 2>nul | find "uvicorn" > nul 2>&1
    if !ERRORLEVEL! EQU 0 (
        echo  [STOP] Duke mbylle uvicorn PID: %%a
        taskkill /F /PID %%a > nul 2>&1
        set FOUND=1
    )
)

if "%FOUND%"=="1" (
    echo.
    echo  [OK] Te gjithe proceset u ndalen me sukses.
) else (
    echo.
    echo  [INFO] Asnje server aktiv nuk u gjend ne portat 5050/5051.
)

:: Verifikimi final
timeout /t 1 /nobreak > nul
netstat -an | find ":5050 " | find "LISTENING" > nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [KUJDES] Porta 5050 eshte ende e zene. Provo manualisht: taskkill /F /IM python.exe
) else (
    echo  [OK] Porta 5050 eshte e lire.
)

echo.
echo  ============================================================
echo.
timeout /t 3 /nobreak > nul
