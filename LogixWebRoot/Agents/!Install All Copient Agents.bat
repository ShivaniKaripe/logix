@echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)
echo ***************************
echo * Copient Agent Installer *
echo ***************************
echo. 
rem Stop all of the Copient Agents - if they are running
echo on
call stop_all.bat

if "%LogixInstallPath%"=="" GOTO NoPath
cd %LogixInstallPath%
cd Agents
:RunInstall
echo Please provide a username and password that the Copient Agents will use 
echo for authentication.
echo.
echo The username should be in the format: 
echo     DomainName\UserName
echo For a local user account use:
echo     .\UserName
echo.
set /P %cptun=Enter the service account username = 
echo.
echo A password that includes special characters needs to be quoted.
set /P %cptpass=Enter the service account password = 
rem set %cptun=NT AUTHORITY\NetworkService
rem set %cptpass=A
set /P %autostart=Set the Agents to start automatically at boot time? (Y/N) = 
set /P %startnow=Start the Agents now? (Y/N) = 

rem Stop all of the Copient Agents - if they are running
echo on
call stop_all.bat
IF NOT %ERRORLEVEL%==0 exit /b 1

@echo off
rem Uninstall any currently installed agents
echo on
call uninstall_All.bat
IF NOT %ERRORLEVEL%==0 exit /b 1


@echo off
echo.

SETLOCAL

SET STARTUP_ARG=Y
IF "%autostart%"=="N" SET STARTUP_ARG=N
IF "%autostart%"=="n" SET STARTUP_ARG=N

rem Install Copient Agents

call install_all.bat "%cptun%" "%cptpass%" %STARTUP_ARG%
IF NOT %ERRORLEVEL%==0 exit /b 1

ENDLOCAL

echo.


@echo off


:CheckStart
@echo off
if "%startnow%"=="Y" GOTO StartAgents
if "%startnow%"=="y" GOTO StartAgents
GOTO FinishUp
:StartAgents
rem Start the Copient Agents
echo on
call start_all.bat
IF NOT %ERRORLEVEL%==0 exit /b 1


:FinishUp
@echo off
set %cptun=
set %cptpass=
set %startnow=
set %autostart=
if not "%1"=="/savelog" erase AgentInstall.log
echo.
echo ************************************
echo * Copient Agent Installer Complete *
echo ************************************
echo.
GOTO End

:NoPath
Echo The environment variable LogixInstallPath is not set!
GOTO End

:End
pause


