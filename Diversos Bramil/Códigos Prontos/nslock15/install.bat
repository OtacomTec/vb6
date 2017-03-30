@echo off

IF EXIST %WinDir%\system\nslock15vb5.ocx goto ERRO
IF EXIST %WinDir%\system\nslock15vb6.ocx goto ERRO

:MAIN

cls
echo This will install ActiveLock on your system
echo Press Ctrl+C to abort, or
pause

echo.
echo Copying files...
copy nslock15vb5.ocx %WinDir%\system
copy nslock15vb6.ocx %WinDir%\system

echo.
echo Registering files...
cd %WinDir%\system
regsvr32/s nslock15vb5.ocx
regsvr32/s nslock15vb6.ocx

echo.
echo ActiveLock was just installed in your system!
echo.
echo Now, open VB and enable ActiveLock control in the 
echo Controls tab of the Components dialog box. 

goto END

:ERRO

echo ActiveLock is already installed on your system.
echo.

goto END

:END
