set ERROR_LOG=d:\error\build.log
set TODIR=c:\inetpub\fnsnet\
set REG=regsvr32 
set UNREG=regsvr32 /u 

	echo START Moving Objects >  %ERROR_LOG%
	net stop "Microsoft SMTP Service"
	net stop "World Wide Web Publishing Service"
	net stop "FTP Publishing Service"
	net stop "IIS Admin Service"
	
	Attrib -R %TODIR%\*.* /S

	echo %UNREG% Objects >>  %ERROR_LOG%
	%UNREG% "%TODIR%bin\FNSUtils.dll"
	%UNREG% "%TODIR%bin\BusinessObjects.dll.dll"
	%UNREG% "%TODIR%bin\PersistenceBroker.dll"
	%UNREG% "%TODIR%bin\CallObject.dll"

	%UNREG% "%TODIR%bin\VBSecurity.dll"
	%UNREG% "%TODIR%bin\Eval.dll" 		
	%UNREG% "%TODIR%bin\VBAddressBook.dll" 
	%UNREG% "%TODIR%bin\AHSearch.dll"
	%UNREG% "%TODIR%bin\BranchAssignment.dll"
	%UNREG% "%TODIR%bin\MCBranchAssignment.dll"
	%UNREG% "%TODIR%bin\ReportCall.dll"

	%UNREG% "%TODIR%bin\faxtran.dll" 		
	%UNREG% "%TODIR%bin\PrintIt.dll"
	%UNREG% "%TODIR%bin\ICMSRouting.dll
	%UNREG% "%TODIR%bin\ObjectFactory.dll
	%UNREG% "%TODIR%bin\OutboundEDIComponent.dll"
	%UNREG% "%TODIR%bin\OutboundEDIMoverCom.dll"
 	%UNREG% "%TODIR%bin\PrintBatchManager.dll
 	%UNREG% "%TODIR%bin\ResolveCallVars.dll

	echo %REG% Objects >>  %ERROR_LOG%
	%REG% "%TODIR%bin\FNSUtils.dll"
	%REG% "%TODIR%bin\BusinessObjects.dll.dll"
	%REG% "%TODIR%bin\PersistenceBroker.dll"
	%REG% "%TODIR%bin\CallObject.dll"

	%REG% "%TODIR%bin\VBSecurity.dll"
	%REG% "%TODIR%bin\Eval.dll" 		
	%REG% "%TODIR%bin\VBAddressBook.dll" 
	%REG% "%TODIR%bin\AHSearch.dll"
	%REG% "%TODIR%bin\BranchAssignment.dll"
	%REG% "%TODIR%bin\MCBranchAssignment.dll"
	%REG% "%TODIR%bin\ReportCall.dll"

	%REG% "%TODIR%bin\faxtran.dll" 		
	%REG% "%TODIR%bin\PrintIt.dll"
	%REG% "%TODIR%bin\ICMSRouting.dll
	%REG% "%TODIR%bin\ObjectFactory.dll
	%REG% "%TODIR%bin\OutboundEDIComponent.dll"
	%REG% "%TODIR%bin\OutboundEDIMoverCom.dll"
 	%REG% "%TODIR%bin\PrintBatchManager.dll
 	%REG% "%TODIR%bin\ResolveCallVars.dll

	net start "IIS Admin Service"  >>  %ERROR_LOG%
	net start "World Wide Web Publishing Service"  >>  %ERROR_LOG%

	echo DONE Moving Objects >>  %ERROR_LOG%
	notepad %ERROR_LOG%

