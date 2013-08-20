set ERROR_LOG=d:\fnsdata\error\build.log
set TODIR=d:\designer\
set REG=regsvr32


	echo START Moving Objects >  %ERROR_LOG%
	net stop "IIS Admin Service" /Y
	
	Attrib -R %TODIR%\*.*

	echo %REG% Objects >>  %ERROR_LOG%
	%REG% "%TODIR%lib\FNSDSecurity.dll"
	%REG% "%TODIR%lib\FNSDClipboard.dll"
	%REG% "%TODIR%lib\FNSDMigrationScheduler.dll"
	%REG% "%TODIR%lib\FNSPageEditor.dll"	
	%REG% "%TODIR%reports\excelclass.dll"
	
	

	echo DONE Moving Objects >>  %ERROR_LOG%
	notepad %ERROR_LOG%

