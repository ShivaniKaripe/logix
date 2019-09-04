@echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)
REM  Uninstall a single agent.

SETLOCAL
SET DATESTAMP=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
SET TIMEHOUR=%TIME:~0,2%
SET TIMESTAMP=%TIMEHOUR: =0%.%TIME:~3,2%.%TIME:~6,2%%

REM  Params
REM    Agent exe

SET PARAM_EXE=%~1
SHIFT
SET PARAM_DOTNETPATH=%~1
SHIFT

IF "%PARAM_EXE%"=="" GOTO Usage

IF NOT EXIST "%PARAM_EXE%" (
    SET ERRMSG=Executable "%PARAM_EXE%" does not exist
    goto ErrorExit
)

IF "%PARAM_DOTNETPATH%"=="" SET PARAM_DOTNETPATH="C:\Windows\Microsoft.NET\Framework\v4.0.30319"

REM -- Ignore failures here. If the agent is already uninstalled, we're fine.
REM -- If there's another reason that we might fail, we might want to investigate later.
echo -- %DATESTAMP%_%TIMESTAMP% %PARAM_DOTNETPATH%\installutil.exe /logfile=AgentUninstall.log /LogToConsole=false /U "%PARAM_EXE%" -- >>AgentUninstall.log
%PARAM_DOTNETPATH%\installutil.exe /logfile=AgentUninstall.log /LogToConsole=false /U "%PARAM_EXE%" >NUL 2>&1

SET SERVICE_FILE=%PARAM_EXE:.exe=.Service%
REM echo SERVICE_FILE: %SERVICE_FILE%

REM  Ignore failure here. If the file is already gone, it's not a problem.
del /Q %SERVICE_FILE% 2>NUL

ENDLOCAL
exit /b 0


:Usage

echo Usage: uninstall_agent.bat ^<agent_exe^> 1>&2
echo     agent_exe - Executable name 1>&2

ENDLOCAL
exit /b 1

:ErrorExit

echo Error: %ERRMSG% 1>&2

ENDLOCAL
exit /b 1
