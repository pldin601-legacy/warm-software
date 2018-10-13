@echo off
echo Checking...
if exist %windir%\system32\comctl32.ocx goto end
echo Registering...
copy comctrl32.ocx %windir%\system32\
regsvr32 /u /s comctl32.ocx
regsvr32 /s comctl32.ocx
echo OK.
goto pend

:end
echo Control is allready registered!

:pend
pause