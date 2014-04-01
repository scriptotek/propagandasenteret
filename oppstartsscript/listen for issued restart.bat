@echo off
echo %date% %time% Sjekker etter restartkommando...

REM DEFAULTS
cd\show\script
echo 0 > restartmaskin.txt

REM HER STARTER LOOPEN
:startListen

set /p IssueRestart=<restartmaskin.txt

REM SJEKKER ETTER RESTART-KOMMANDO
IF %IssueRestart% == 1 (

echo. && echo.
echo Restart iverksatt!
shutdown.exe /r /t 0 /f

EXIT

)

timeout /t 5 >NUL
goto StartListen