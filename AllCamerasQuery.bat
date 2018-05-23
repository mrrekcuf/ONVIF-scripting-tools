echo off
set csv=%~dp0IP Camera Logins - Sheet1.csv
rem type "%csv%"

rem SETLOCAL ENABLEDELAYEDEXPANSION
setlocal enabledelayedexpansion
for /f "usebackq  tokens=1,2,3,4,5,6,7,8,9 delims=," %%i in ("%csv%") do (
 	
	set ip=%%j 
	set user=%%o
	set pw=%%p
	if not "%%k"=="--" set ip=%%j:%%k
	@echo %%i !ip! !user! !pw!
	cscript /nologo "%~dp0onvifQuery.vbs" !ip! !user! !pw! 
	@echo.
)

rem pause