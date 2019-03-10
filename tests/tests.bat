title SilentUninstaller Test
@echo off
@cd /d "%~dp0"

call :def_main
cls & echo Ok!
pause

:def_main
	set Branch="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	set BranchNode="DisplayName"
	set NodeValue="Mozilla Maintenance Service"
	set NodeForLook="DisplayVersion"
	set NodeForLookConstValue="8.3.13.1513"
	
	@cd ..
	call "%cd%\uninstall.vbs" %Branch% %BranchNode% %NodeValue% %NodeForLook% %NodeForLookConstValue% %SilentKey%
	exit /b