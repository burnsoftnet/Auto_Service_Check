@ECHO OFF
SET APPPATH=C:\scripts
SET APPLOG=%APPPATH%\logs
SET VBScript=Auto_Service_Check
SET MAXSIZE=700000
SET LOGFILE=%APPLOG%\%VBScript%.log

IF EXIST "%LOGFILE%" (
	FOR /F "usebackq" %%A IN ('%LOGFILE%') DO set LOGSIZE=%%~zA

	if %LOGSIZE% GTR %MAXSIZE% (
		echo.File is ^>= %MAXSIZE% bytes
		echo "delete logfile"
		del /f /q %LOGFILE%
	) ELSE (
		echo.File is ^< %MAXSIZE% bytes
	)
)

c:\windows\SysWOW64\cscript /nologo %APPPATH%\%VBScript%.vbs /interval=5 >> %LOGFILE%
