@echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)
REM  Install a single agent.


SETLOCAL
SET DATESTAMP=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
SET TIMEHOUR=%TIME:~0,2%
SET TIMESTAMP=%TIMEHOUR: =0%.%TIME:~3,2%.%TIME:~6,2%%

REM  Params
REM    Agent exe
REM    Username
REM    Password
REM    Start now?
REM    Start at startup?
REM    .NET Path to installutil.exe


SET PARAM_EXE=%~1
SHIFT
SET PARAM_USER=%~1
SHIFT
SET PARAM_PASS="%~1"
SHIFT
SET PARAM_STARTUP=%~1
SHIFT
SET PARAM_DOTNETPATH=%~1
SHIFT

SET SERVICE_FILE=%PARAM_EXE:.exe=.Service%

IF %PARAM_PASS%=="" GOTO Usage
IF "%PARAM_DOTNETPATH%"=="" SET PARAM_DOTNETPATH="C:\Windows\Microsoft.NET\Framework\v4.0.30319"

REM  Sanity checks

REM  Does the executable exist?
IF NOT EXIST "%PARAM_EXE%" (
    echo Error: Could not find agent file "%PARAM_EXE%" 1>&2
    GOTO ErrorExit
)


REM  Remove the leading .\ since most of the subsequent tools don't agree with it.
SET ADJUSTED_USER=%PARAM_USER:.\=%
REM echo ADJUSTED_USER: %ADJUSTED_USER%

REM Check does not work with network users
REM  Does the user exist?
REM net user %ADJUSTED_USER% >NUL 2>&1
REM IF NOT %ERRORLEVEL%==0 (
REM    echo Error: User '%PARAM_USER%' does not exist on this system 1>&2
REM   GOTO ErrorExit
REM )


REM  Does the user/pass combo work?
REM  @todone: Find out if this is something that we can rely on everywhere.
REM    No, we can't. This blew up on Huw's test VM. Need to find another
REM    way to verify user password.
REM net use \\127.0.0.1 /USER:%ADJUSTED_USER% %PARAM_PASS% >NUL 2>&1
REM IF NOT %ERRORLEVEL%==0 (
REM     echo Error: Incorrect password provided for user '%PARAM_USER%' 1>&2
REM     GOTO ErrorExit
REM )

REM  Cleanup the connection that we just created.
REM net use /DELETE \\127.0.0.1 >NUL 2>&1


IF "%PARAM_STARTUP%"=="" GOTO ParamCheckDone

IF %PARAM_STARTUP%==Y GOTO ParamCheckDone
IF %PARAM_STARTUP%==y GOTO ParamCheckDone
IF %PARAM_STARTUP%==N GOTO ParamCheckDone
IF %PARAM_STARTUP%==n GOTO ParamCheckDone

echo Error: Unrecognized value for ^<start_at_startup^?^>: %PARAM_STARTUP% 1>&2
GOTO Usage


:ParamCheckDone




REM  If the .Service file exists, uninstall the agent.
REM    Login credentials may have changed or something, installation should be 
REM    infrequent enough that this isn't too much of a problem.
IF EXIST %SERVICE_FILE% (
    uninstall_agent.bat "%PARAM_EXE%" %PARAM_DOTNETPATH%
)


REM -- Attempt to install the agent.
echo -- %DATESTAMP%_%TIMESTAMP% %PARAM_DOTNETPATH%\installutil.exe /logfile=AgentInstall.log /username="%PARAM_USER%" /password=* "%PARAM_EXE%" -- >>AgentInstall.log
%PARAM_DOTNETPATH%\installutil.exe /logfile=AgentInstall.log /username="%PARAM_USER%" /password=%PARAM_PASS% "%PARAM_EXE%" >agent_install_out.txt 2>agent_install_err.txt
IF NOT %ERRORLEVEL%==0 (
    echo Error: Problem trying to install "%PARAM_EXE%" 1>&2
    type agent_install_out.txt 1>&2
    type agent_install_err.txt 1>&2
    GOTO ErrorExit
)

DEL /Q agent_install_err.txt >NUL 2>&1

SET SERVICE_NAME=
REM -- Store a file with the agent service name for future reference.
FOR /F "tokens=2" %%A IN ('FINDSTR /C:" has been successfully installed." agent_install_out.txt') DO SET SERVICE_NAME=%%A

DEL /Q agent_install_out.txt >NUL 2>&1

IF "%SERVICE_NAME%"=="" (
    echo Error: No services found in '%PARAM_EXE%' 1>&2
    GOTO ErrorExit
)


REM SET SERVICE_NAME=%SERVICE_NAME:~8%
REM SET SERVICE_NAME=%SERVICE_NAME: has been successfully installed.=%
REM echo SERVICE_NAME: %SERVICE_NAME%

echo %SERVICE_NAME% > %SERVICE_FILE%



REM -- Set the service to manual start if this was requested.
IF "%PARAM_STARTUP%"=="" GOTO SetStartupDelayedStart
IF %PARAM_STARTUP%==N GOTO SetStartupManual
IF %PARAM_STARTUP%==n GOTO SetStartupManual

REM If %PARAM_STARTUP%==Y or y then set service to Automatic (Delayed Start) to allow all these services that are dependencies of AMS agents to start first.
:SetStartupDelayedStart
echo -- sc config %SERVICE_NAME% start= delayed-auto -- >>AgentInstall.log
sc config %SERVICE_NAME% start= delayed-auto >>AgentInstall.log 2>&1
GOTO Done

REM If %PARAM_STARTUP%==N or n then set service to Manual so it does not start up when user reboots their machine.
:SetStartupManual
echo -- sc config %SERVICE_NAME% start= demand -- >>AgentInstall.log
sc config %SERVICE_NAME% start= demand >>AgentInstall.log 2>&1
GOTO Done

:Done

ENDLOCAL
exit /b 0


:Usage

echo Usage: install_agent.bat ^<agent_exe^> ^<username^> ^<password^> [ ^<start_at_startup^?^> ] 1>&2
echo     agent_exe - Executable name 1>&2
echo     username - Windows account under which the agent will run 1>&2
echo     password - Password for the Windows account under which the agent will run 1>&2
echo                A password that includes special characters needs to be quoted. 1>&2
echo     start_at_startup - (y^|Y^|n^|N) (Default "Y") 1>&2
echo                YES = Set the Windows service to "Automatic (Delayed Start)" to allow all the services that are dependencies to start first. 1>&2
echo                NO  = Set the Windows service to "Manual" so it does not start up when the machine is booted or rebooted. 1>&2

:ErrorExit

ENDLOCAL

exit /b 1





ENDLOCAL
