@Echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)

cls
if %1x==x then goto nothingtodo

echo Removing previously %1 DLL...
gacutil /u %1

echo Installing %1 DLL...
gacutil /i "..\Libraries\%1.dll"
IF ERRORLEVEL 1 "ERROR:  gacutil /i ..\Libraries\%1.dll failed to register."

echo Reseting IIS
iisreset

goto end

:nothingtodo
echo Done!
Pause

:end