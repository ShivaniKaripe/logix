@echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)
REM  Start an agent. Install it if it needs to be and is possible.


REM  Params
REM    Agent exe
REM    User
REM    Pass


SETLOCAL
SET DATESTAMP=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
SET TIMEHOUR=%TIME:~0,2%
SET TIMESTAMP=%TIMEHOUR: =0%.%TIME:~3,2%.%TIME:~6,2%%

SET PARAM_EXE=%~1
SHIFT
SET PARAM_USER=%~1
SHIFT
SET PARAM_PASS=%~1
SHIFT

IF "%PARAM_EXE%"=="" GOTO Usage

REM  Sanity check
IF NOT EXIST "%PARAM_EXE%" (
    SET ERRMSG=Could not find agent file "%PARAM_EXE%"
    GOTO ErrorExit
)


SET SERVICE_FILE=%PARAM_EXE:.exe=.Service%

IF EXIST %SERVICE_FILE% goto StartAgent

IF "%PARAM_PASS%"=="" (
    SET ERRMSG=The agent "%PARAM_EXE%" hasn't been installed and no credentials have been provided for installation.
    GOTO ErrorExit
)

CALL install_agent.bat "%PARAM_EXE%" "%PARAM_USER%" "%PARAM_PASS%"
REM  We assume that the above call wrote to the error log.
IF NOT %ERRORLEVEL%==0 (
    SET ERRMSG=Could not start the agent "%PARAM_EXE%"
    GOTO ErrorExit
)

IF NOT EXIST %SERVICE_FILE% (
    SET ERRMSG=Agent installed, but service file "%SERVICE_FILE%" could not be found.
    GOTO ErrorExit
)

:StartAgent

SET SERVICE_NAME=
FOR /F "tokens=*" %%A in ('type %SERVICE_FILE%') DO SET SERVICE_NAME=%%A

REM echo SERVICE_NAME: %SERVICE_NAME%

echo -- %DATESTAMP%_%TIMESTAMP% net start %SERVICE_NAME% -- >> AgentStart.log
net start %SERVICE_NAME% > agent_start_out.txt 2>&1

SET TASK_RESULT=%ERRORLEVEL%

type agent_start_out.txt >> AgentStart.log

IF %TASK_RESULT%==0 GOTO Done


REM  If the service was already started, we're fine. Otherwise, it's a problem.
SET ERROR_STATE=1
FOR /F %%A IN ('FINDSTR /C:"The requested service has already been started." agent_start_out.txt') DO SET ERROR_STATE=0

IF %ERROR_STATE%==1 (
    echo Error: '%SERVICE_NAME%' could not be started for the following reason: 1>&2
    type agent_start_out.txt 1>&2
    GOTO ErrorNoMsgExit
)


:Done

DEL /Q agent_start_out.txt > NUL 2>&1

ENDLOCAL
exit /b 0


:Usage

echo Usage: start_agent.bat ^<agent_exe^> [ ^<username^> ^<password^> ] 1>&2
echo     agent_exe - Executable name 1>&2
echo     username - Windows account under which the agent will run 1>&2
echo     password - Password for the Windows account under which the agent will run 1>&2

ENDLOCAL
exit /b 1


:ErrorExit

echo Error: %ERRMSG% 1>&2

:ErrorNoMsgExit

ENDLOCAL
exit /b 1