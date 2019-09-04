@Echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)


SETLOCAL

SET JOB=uninstall_agent.bat
SET INDIVIDUAL_ERROR=uninstall_all0_err.tmp
SET GROUP_ERROR=uninstall_all_err.tmp
SET MESSAGE=Uninstalling

SET PARAM_DOTNETPATH=%~1

IF "%PARAM_DOTNETPATH%"=="" SET PARAM_DOTNETPATH="C:\Windows\Microsoft.NET\Framework\v4.0.30319"

SET AGENT_LIST_SOURCE=agent_list.txt

IF NOT EXIST %AGENT_LIST_SOURCE% (
    echo Error: 'agent_list.txt' was not found 1>&2
    GOTO ErrorExit
)


REM  Maybe someday we'll go back to installing all exes in a directory.
REM IF EXIST %AGENT_LIST_SOURCE% GOTO DoWork
REM SET AGENT_LIST_SOURCE='dir /B *.exe'


:DoWork

REM -- Accumuate errors in a file so that we can check it later for failure.

DEL /Q %GROUP_ERROR% >NUL 2>&1

SET FoundError=0
FOR /F %%A IN (%AGENT_LIST_SOURCE%) DO (
    echo %MESSAGE% %%A...
    CALL %JOB% %%A %PARAM_DOTNETPATH% 2>%INDIVIDUAL_ERROR%
    type %INDIVIDUAL_ERROR% >>%GROUP_ERROR%
    type %INDIVIDUAL_ERROR% 1>&2
    DEL /Q %INDIVIDUAL_ERROR% >NUL 2>&1
)


REM -- Check for a non-zero error output (ugh).
SET FoundError=0
for %%R in (%GROUP_ERROR%) do if %%~zR gtr 0 SET FoundError=1

IF %FoundError%==1 (
    echo There were errors 1>&2
    GOTO ErrorExit
)

DEL /Q %GROUP_ERROR% >NUL 2>&1

ENDLOCAL
exit /b 0

:ErrorExit

DEL /Q %GROUP_ERROR% >NUL 2>&1

ENDLOCAL
exit /b 1
