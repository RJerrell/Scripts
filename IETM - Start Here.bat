@ECHO off

REM ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
REM + This .bat file removes the LiveContent temp folder structure (\var) from +
REM + the C:\Temp folder that is created during execution of the IETM.  This   +
REM + has been turned into the supplier as needing a permanent fix.	       +
REM ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

RMDIR /S /Q C:\Temp\var

.\autoplay_cdonly.exe