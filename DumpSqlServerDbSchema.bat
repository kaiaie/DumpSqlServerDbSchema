@ECHO OFF
:: \brief Batch file wrapper for Dump SQL Server schema script
:: \author Ken Keenan, <mailto:ken@kaia.ie>

SET myname=%~0
:: Add extension if necessary
IF NOT "%myname:~-4%" == ".bat" GOTO addbat
IF NOT "%myname:~-4%" == ".BAT" GOTO addbat
GOTO checkcd

:addbat
SET myname=%myname%.bat

:checkcd
SET mydir=
:: If file exists, it could be in the current directory or a relative path
IF NOT EXIST %myname% GOTO checkpath
:: Replace backslashes with underscores; if no change then not a relative path
SET tmpfile = %myname:\=_%
IF "%myname%" == "%tmpfile%" GOTO setcd
FOR %%i IN ("%myname%") DO SET mydir=%%~dpi
GOTO gotdir

:setcd
:: In current directory
SET mydir=%CD%\
GOTO gotdir

:checkpath
:: Check the PATH environment variable for the batch file
SET fullname=
FOR %%i IN (%myname%) DO SET fullname=%%~$PATH:i
IF NOT "%fullname%" == "" FOR %%i IN ("%fullname%") DO SET mydir=%%~dpi
GOTO gotdir

:gotdir
SET scriptfile=%mydir%DumpSqlServerDbSchema.vbs
IF NOT EXIST "%scriptfile%" GOTO err
cscript //nologo "%scriptfile%" %*
GOTO done

:err
ECHO Cannot find script file "%scriptfile%"
GOTO done

:done
ECHO.

