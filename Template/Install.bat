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
REM VERSION			8.7.5.0
REM FILENAME		notepad++.exe
REM OWNER			Thomas Hoins
REM ASSETNUMBER		TOM-0061



REM install Notepad++
"%SOURCEDIR%\npp.8.7.5.Installer.x64.exe" /S /noUpdater
"%windir%\system32\xcopy.exe" "%SOURCEDIR%\filestocopy\*" "C:\Program Files\" /E /H /I /Q /Y
