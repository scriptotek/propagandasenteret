@ECHO OFF
mode con:cols=140 lines=60

SET revision=1
SET "title=%~nx0 - revision %version%"
TITLE %title%

:StartListen
CD \show\script
cscript //Nologo infoskjerm_controller.vbs
ECHO infoskjerm_controller.vbs avsluttet. Starter på nytt om 10 sekunder
timeout /t 10 > NUL
GOTO StartListen


EXIT
