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
REM VERSION			8.9.4.0
REM FILENAME		notepad++.exe
REM OWNER			M. Marx
REM ASSETNUMBER		A4L-0061



REM install Notepad++
"%windir%\system32\msiexec.exe" /i "%sourcedir%\npp.8.9.1.Installer.x64.msi" /quiet /norestart ALLUSERS=2 /l "%logpath%\Don Ho_Notepad++_8.9.1_MUI.log"
"%windir%\system32\xcopy.exe" "%SOURCEDIR%\filestocopy\*" "C:\Program Files\" /E /H /I /Q /Y
