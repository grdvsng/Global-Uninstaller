title SilentUninstaller
@echo off
@cd/d "%~dp0"

call :IsAdmin
call :Run
exit


:Run
	set /p argv0="Pass the registry branch with uninstall data. > "
	set /p argv1="Pass the registry node with special application value (name or something else). > "
	set /p argv2="Pass the value that NodeReg must be. > "
	set /p argv3="Pass the node that have special parameter an if app get this, script must skip this version program for uninstall. > "
	set /p argv4="Pass the value that constant node must have. > "
	
	call "uninstall.vbs" %argv0% %argv1% %argv2% %argv3% %argv4%
	
	timeout /t 3 /Nobreak >nul
	exit /b
	
:IsAdmin 
	AT > nul
	if errorlevel 1 (
		cls & echo Need Administrator access.
		pause
		exit
	) else (
		cls & echo start ^> uninstall.vbs
		exit /b
	)