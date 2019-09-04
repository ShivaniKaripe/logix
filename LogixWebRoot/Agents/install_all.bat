@echo off
REM version:7.3.1.138972.Official Build (SUSDAY10202)


SETLOCAL

REM -- Params
REM --   User
REM --   Pass
REM --   Startup
REM --   .NET path to installutil.exe


SET JOB=install_agent.bat
SET INDIVIDUAL_ERROR=install_all0_err.tmp
SET GROUP_ERROR=install_all_err.tmp
SET MESSAGE=Installing



SET PARAM_USER=%~1
SHIFT
SET PARAM_PASS="%~1"
SHIFT
SET PARAM_STARTUP=%~1
SHIFT
SET PARAM_DOTNETPATH=%~1
SHIFT

IF %PARAM_PASS%=="" GOTO Usage

IF "%PARAM_STARTUP%"=="" SET PARAM_STARTUP=Y

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
    CALL %JOB% %%A %PARAM_USER% %PARAM_PASS% %PARAM_STARTUP% %PARAM_DOTNETPATH% 2>%INDIVIDUAL_ERROR%
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


:Usage

echo Usage: install_All.bat ^<username^> ^<password^> [ ^<start_at_startup^?^> ] 1>&2
echo     username - Windows account under which the agent will run 1>&2
echo     password - Password for the Windows account under which the agent will run 1>&2
echo                A password that includes special characters needs to be quoted. 1>&2
echo     start_at_startup - (y^|Y^|n^|N) (Default "Y") 1>&2
echo                YES = Set the Windows service to "Automatic (Delayed Start)" to allow all the services that are dependencies to start first. 1>&2
echo                NO  = Set the Windows service to "Manual" so it does not start up when the machine is booted or rebooted. 1>&2

exit /b 1


:ErrorExit


DEL /Q %GROUP_ERROR% >NUL 2>&1

ENDLOCAL
exit /b 1



REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "AirMileRejectionAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPEOfferAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-GM.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-PA.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-PA-parallel.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-RA-N.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-RA-ND.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-RA-OD.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-RD.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-SF.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-SV.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-UL.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-UR.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CPETransUpdateAgent-YB.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CRMExportAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CRMImportAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CustomerRemovalAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "CustomerUpdateAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "DataExportAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "DBPurgeAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "HouseholdUpdateAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "IssuanceDBPurge.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "IssuanceExtract.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "LocationHierarchyAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "LocationUpdateAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "OfferCustomerAgentMT.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "OfferFileAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "OfferValidationAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "PointsHistoryMovementAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ProcessCustomerGroups.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ProcessIssuance.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ProcessPointsPrograms.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ProcessProductGroups.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ProductHierarchyAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ProductUpdateAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "PromoMovementAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "ReportingAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "TrafficCopAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "TransHistoryMovementAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "TransRedemptionMovementAgent.exe"
REM installutil /logfile=AgentInstall.log /username="%cptun%" /password="%cptpass%" "WatchDog.exe"

