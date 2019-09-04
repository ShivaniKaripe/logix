@ECHO OFF
CALL SetBatVariables.bat

REM AMS- Migrating Attributes to Preferences

SETLOCAL
cd msg3_scripts

IF "%DBS%" ==""  GOTO No1
IF "%UName%" ==""  GOTO No1
IF "%LRT%" ==""  GOTO No1
IF "%LXS%" ==""  GOTO No1
IF "%Pwd%" ==""  GOTO No1
IF "%PMRT%" ==""  GOTO No1

::echo %1%2
::echo %DBS%
::echo %UName%
::echo %PWd%
::echo %LRT%
::echo %PMRT%
::pause 

echo ***************************
echo * Executing SQL scripts *
echo ***************************
echo. 

if not "%SQLPort%" =="" GOTO USEPORT
echo Executing CPEAttributesToPrefs.sql....
sqlcmd -S%DBS% -U%UName% -P%Pwd% -d%PMRT% -iCPEAttributesToPrefs.sql -b
if %ERRORLEVEL% EQU 0 GOTO Query1


:Query1
echo Executing CPEAttributesToPrefs.sql....
sqlcmd -S%DBS% -U%UName% -P%Pwd% -d%PMRT%  -iCPEAttributesToPrefs.sql -b
if NOT %ERRORLEVEL% EQU 0 (GOTO END1) 


GOTO END

:No1
  ECHO Missing parameters
  ECHO Configure batchConfig.cfg file with current environment

GOTO END


:END

ENDLOCAL

echo. 
echo ************************************************
echo * Successful - fininshed Executing SQL scripts *
echo ************************************************
echo.
GOTO END2
@Echo OFF

:END1
echo. 
echo ************************************************************
echo * Unsucessful - Error occured while executing SQL scripts. *
echo ************************************************************
echo.
ENDLOCAL
@Echo OFF

:END2