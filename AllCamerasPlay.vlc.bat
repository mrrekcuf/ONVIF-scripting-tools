echo off
set csv=%~dp0IP Camera Logins - Sheet1.csv
rem type "%csv%"

setlocal enabledelayedexpansion
for /f "usebackq  tokens=1,2,3,4,5,6,7,8,9 delims=," %%i in ("%csv%") do (
 	
	set ip=%%j 
	set user=%%o
	set pw=%%p
	if not "%%k"=="--" set ip=%%j:%%k
	@echo %%i !ip! !user! !pw!
	for /f "usebackq   tokens=1,2,3,4,5,6,7,8 delims=: " %%a in (`cscript /nologo "%~dp0onvifQuery.vbs" !ip! !user! !pw! ^| findstr "Name"`) do (
		setlocal enabledelayedexpansion
		rem echo %%a %%b %%c %%d
		rem echo !ip! !user! !pw! %%d
		cscript /nologo "%~dp0onvifPlay.vlc.vbs" !ip! !user! !pw! %%d

	)
	@echo.
)

rem pause