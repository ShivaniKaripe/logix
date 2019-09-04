::This .bat file reads user defined .bat file variables
::configuration from batchConfig.cfg file and sets them.

@Echo Off
::SETlocal EnableDelayedExpansion
SET DBS=
SET SQLPort=
SET UName=
SET Pwd=
SET LRT=
SET LXS=
SET PMRT=
::SET ELC=

for /f "tokens=*" %%a in (batchConfig.cfg) do call :processLine %%a
goto :eof


:processLine
::echo First token: %1, second token: %2
if "%1"=="DBServer" set DBS=%2
if "%1"=="SQLPortNo" set SQLPort=%2
if "%1"=="DBUName" set UName=%2
if "%1"=="DBPwd" set Pwd=%2
if "%1"=="DBLRT" set LRT=%2
if "%1"=="DBLXS" set LXS=%2
if "%1"=="DBPMRT" set PMRT=%2
::echo %DBS%,
::echo %SQLPort%,
::echo %UName%,
::echo %Pwd%,
::echo %LRT%,
::echo %LXS%,
::echo %PMRT%
goto :eof


:eof