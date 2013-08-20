<%
Function CheckLogicExpression(strExpression, strLanguage)
    Dim strCode
    
    strCode = "Sub TestFunc()" & vbCrLf & "If (" & strExpression & ") Then " & vbCrLf & "End If" & vbCrLf & "End Sub"
    On Error Resume Next
    
    Set sc = Session("ScriptControl")
    sc.language = strLanguage
    sc.AddCode strCode
    sc.Run "TestFunc"

	If Err > 0 Then
		CheckLogicExpression = False
		Exit Function
	End If
	
	CheckLogicExpression = True
	Set sc = Nothing
End Function
%>