<%
function CheckADOErrors(curConn, strContext)
	Dim intLoop, bFound, strError, strComponent
	strError = ""
	bFound = False
	intLoop = 0 
	On Error Resume Next
	If curConn.Errors.count > 0 Then
		bFound = True
		strError = "Encountered errors in " & strContext & ":<br> "
		For intLoop = 0 To curConn.Errors.Count -1
			strError = strError & "Error Number " & curConn.Errors(intLoop).Number
			strError = strError & " - " & curConn.Errors(intLoop).Description & "<P>"
		Next
	End If
	curConn.Errors.Clear()
	CheckADOErrors = strError
end function
%>