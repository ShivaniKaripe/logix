@echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)
REM -- Stop an agent

REM -- Params
REM --   Agent exe

SETLOCAL
SET DATESTAMP=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
SET TIMEHOUR=%TIME:~0,2%
SET TIMESTAMP=%TIMEHOUR: =0%.%TIME:~3,2%.%TIME:~6,2%%

SET PARAM_EXE=%~1
SHIFT

IF "%PARAM_EXE%"=="" GOTO Usage


SET SERVICE_FILE=%PARAM_EXE:.exe=.Service%


REM -- Sanity check
IF NOT EXIST "%PARAM_EXE%" (
    echo Error: Could not find agent file "%PARAM_EXE%" 1>&2
    GOTO ErrorExit
)


REM -- Nothing to stop. The agent should not have been installed.
IF NOT EXIST "%SERVICE_FILE%" GOTO Done


SET SERVICE_NAME=
FOR /F %%A in ('TYPE "%SERVICE_FILE%"') DO SET SERVICE_NAME=%%A

REM echo SERVICE_NAME: %SERVICE_NAME%

echo -- %DATESTAMP%_%TIMESTAMP% net stop %PARAM_EXE% -- >> AgentStop.log
net stop %SERVICE_NAME% > agent_stop_out.txt 2>&1

SET TASK_RESULT=%ERRORLEVEL%

type agent_stop_out.txt >> AgentStop.log

IF %TASK_RESULT%==0 GOTO Done


SET ERROR_STATE=1
FOR /F %%A IN ('FINDSTR /C:" service is not started." agent_stop_out.txt') DO SET ERROR_STATE=0

IF %ERROR_STATE%==1 (
    echo Error: '%SERVICE_NAME%' could not be stopped for the following reason: 1>&2
    type agent_stop_out.txt 1>&2
    GOTO ErrorExit
)


:Done


DEL /Q agent_stop_out.txt > NUL 2>&1

ENDLOCAL
exit /b 0


:Usage

echo Usage: stop_agent.bat ^<agent_exe^> 1>&2
echo     agent_exe - Executable name 1>&2

:ErrorExit

ENDLOCAL
exit /b 1
