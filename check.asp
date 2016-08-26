<%
'----------------------------------------------------------------------
' FILE: Check.asp for Designer	//	RA	10/03
'----------------------------------------------------------------------

const cLogFilename = "C:\checkErrlog.txt"

dim aActions(7)
dim oConn, oRS, oDBPool, oSec, x, lWithError

aActions(1) = "Creating connection object"
aActions(2) = "Creating recordset object"
aActions(3) = "Running query on USERS table"
aActions(4) = "Moving to the next record on Users table"
aActions(5) = "Running query on AHS table"
aActions(6) = "Moving to the next record on AHS table"
aActions(7) = "Releasing connection object"

' start the state machine
x = 1
on error resume next
lWithError = false
do while not lWithError
	eval "f" & x
	lWithError = checkError( aActions(x) )
	x = x + 1
	if x > ubound(aActions) then
		exit do
	end if
loop	
set oDBPool = nothing
set oConn = nothing
set oRS = nothing
set oSec = nothing
if not lWithError then
	Response.Write("Server tested OK")
end if
session.Abandon
	

function f1
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "DSN=FNSBA;UID=FNSOWNER;PWD=CTOWN_DESIGNER"
end function

function f2
Set oRS = Server.CreateObject("ADODB.Recordset")
end function

function f3
oRS.Open "Select * From USERS Where rownum = 1", oConn
end function

function f4
oRS.MoveNext
end function

function f5
oRS.Close 
oRS.Open "SELECT * FROM ACCOUNT_HIERARCHY_STEP Where rownum = 1", oConn
end function

function f6
oRS.MoveNext
end function

function f7
oConn.Close 
end function

function f8
set oSec = server.createObject("Security.CSecurity")
end function

function checkError(cAction)
checkError = false
If err.number <> 0 then
	logError cAction
	response.write "An error occured. Check log file '" & cLogFilename & "'"
	response.write err.Description 
	checkError = true
end if
end function

sub logError(cAction)
dim oFSO, oFH

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFH = oFSO.OpenTextFile(cLogFilename, 8, true)
oFH.WriteLine(now & "| An error ocurred| ACTION: " & cAction & "| ERROR NO.: " & err.number & "| ERROR TEXT: " & err.Description )
oFH.Close
Set oFSO = nothing
Set oFH = nothing
end sub
%>

