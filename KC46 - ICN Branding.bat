Set cgmSrc=%1
Set cgmDest=%2

REM Remove the REM statements to debug these values passed into this batch file
REM echo %cgmSrc%
REM echo %cgmDest%
REM PAUSE

"C:\Program Files (x86)\PTC\PTC Arbortext IsoDraw 7.3\Program\IsoDraw73.exe" -q -batch -s%cgmSrc% -d%cgmDest% 	-f10 -m"ADD ICN"
exit