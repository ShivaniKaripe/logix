@ECHO OFF
CALL SetBatVariables.bat

REM AMS- Migrating customer attributes to preferences

SETLOCAL
cd msg3_scripts

IF "%DBS%" ==""  GOTO No1
IF "%UName%" ==""  GOTO No1
IF "%LRT%" ==""  GOTO No1
IF "%LXS%" ==""  GOTO No1
IF "%Pwd%" ==""  GOTO No1

::echo %1%2
::echo %DBS%
::echo %UName%
::echo %PWd%
::echo %LRT%
::pause 

echo ***************************
echo * Executing SQL scripts *
echo ***************************
echo. 

if not "%SQLPort%" =="" GOTO USEPORT
echo Executing CustAttributesToPrefs.sql....
sqlcmd -S%DBS% -U%UName% -P%Pwd% -d%LXS% -iCustAttributesToPrefs.sql -b
if %ERRORLEVEL% EQU 0 GOTO Query1


:Query1
echo Executing CustAttributesToPrefs.sql....
sqlcmd -S%DBS% -U%UName% -P%Pwd% -d%LXS%  -iCustAttributesToPrefs.sql -b
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