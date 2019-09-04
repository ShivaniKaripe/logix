@Echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)

if "%LogixInstallPath%"=="" GOTO NoPath

cd %LogixInstallPath%
cd Agents

:RunStart
call start_all.bat
GOTO End

:NoPath
Echo The environment variable LogixInstallPath is not set!
GOTO End

:End
pause

