@echo off
set InstallationsDir=%~dp0
set InstallationsMSI=PhysX-9.13.1220.msi
title %InstallationsDir%
echo.
if not exist "%~dp0%InstallationsMSI%" goto DateiFehlt
rem Prüfe ob Administratorrechte vorhanden sind...
reg add HKEY_LOCAL_MACHINE\Software\Admintest /v "Bei 1 = Adminrechte" /t REG_DWORD /d 00000000 /f >NUL 2>NUL 
reg add HKEY_LOCAL_MACHINE\Software\Admintest /v "Bei 1 = Adminrechte" /t REG_DWORD /d 00000001 /f >NUL 2>NUL 
if errorlevel 1 goto Userlimit
goto Installation
:Installation
echo Installiere %InstallationsMSI%...
start /wait msiexec.exe /i "%InstallationsDir%%InstallationsMSI%" /quiet /norestart /qn
if errorlevel 1 goto Fehler
cls
color 0A
echo.
echo PhysX wurde erfolgreich installiert.
echo.
echo.
PING 1.1.1.1 -n 3 -w 300 >nul
exit

:DateiFehlt
cls
color 0C
echo.
echo "%InstallationsMSI%" fehlt im Ordner "%InstallationsDir%" !
echo.
echo.
PING 1.1.1.1 -n 3 -w 300 >nul
exit

:Userlimit
cls
color 0C
echo.
echo Bitte starten Sie das Batch-Skript mit Administratorrechten!
echo.
echo.
PING 1.1.1.1 -n 3 -w 300 >nul
exit

:Fehler
cls
color 0C
echo.
echo Die Installation von PhysX ist fehlgeschlagen!
echo.
echo.
PING 1.1.1.1 -n 3 -w 300 >nul
exit