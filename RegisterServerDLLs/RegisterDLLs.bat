@Echo off
REM version:6.0.1.84287.Official Build (SUSDAY10083)

SET libraries=..\Libraries\

echo Removing any obsolete libraries ...
for /F "tokens=*" %%f in (obsolete_libs.txt) do (
    IF EXIST ..\Libraries\%%f (
        del ..\Libraries\%%f
    )
    IF EXIST ..\LogixWebRoot\Bin\%%f (
        del ..\LogixWebRoot\Bin\%%f
    )
    IF EXIST ..\LogixWebRoot\App_Code\CSCode\%%f (
        del ..\LogixWebRoot\App_Code\CSCode\%%f
    )
    echo ==>Unregistering %%f
    gacutil /u %%~nf
    echo .
)

echo Removing previously installed DLLs ...

for %%f in ("%libraries%*.dll") do (
    echo ==>Unregistering %%~nf
    gacutil /u %%~nf
    echo .
)

echo Installing DLLs ...

for %%f in ("%libraries%*.dll") do (
    echo ==>Registering ..\Libraries\%%~nxf
    gacutil /i ..\Libraries\%%~nxf
    IF ERRORLEVEL 1 echo "ERROR:  gacutil /i ..\Libraries\%%~nxf failed to register."
    echo .
)

iisreset

echo Done!
