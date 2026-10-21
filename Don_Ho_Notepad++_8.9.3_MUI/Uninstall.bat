@echo off

set MyDir=%~dp0
set SOURCEDIR=%MyDir:~0,-1%
set logpath=%programfiles%\support\logs
IF NOT EXIST "%logPath%" ( mkdir "%logPath%" )

set packagecache=%programfiles%\support\package-cache
IF NOT EXIST "%packagecache%" ( mkdir "%packagecache%" )



REM DESCRIPTION		Notepad++
REM MANUFACTURER	Don Ho
REM LANGUAGE		MUI
REM VERSION			8.9.1.0
REM FILENAME		notepad++.exe
REM OWNER			M. Marx
REM ASSETNUMBER		A4L-0061



REM uninstall Notepad++
"%windir%\system32\msiexec.exe" /x {7349B4F3-02E1-4234-A67A-FA85B33B67AF} /quiet /norestart
rd "%Programfiles%\Notepad++" /s /q
