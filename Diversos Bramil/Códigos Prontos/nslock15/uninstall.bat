@echo off

IF NOT EXIST %WinDir%\system\nslock15vb5.ocx goto ERRO
IF NOT EXIST %WinDir%\system\nslock15vb6.ocx goto ERRO

:MAIN

cls
echo This will uninstall ActiveLock from your system.
echo If you want to quit, press Ctrl+C now, or
pause

echo.
echo IMPORTANT:
echo 1. Make sure that you are not running Visual Basic,
echo 2. Close any application that uses ActiveLock and
pause

echo.
echo Unregistering files...
cd %WinDir%\system
regsvr32/s/u nslock15vb5.ocx
regsvr32/s/u nslock15vb6.ocx

echo.
echo Deleting files...
del nslock15vb5.ocx
del nslock15vb6.ocx

echo.
echo ActiveLock was just uninstalled from your system!
echo.

goto END

:ERRO

echo ActiveLock is not installed on your system.
echo.

goto END

:END