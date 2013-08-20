<%
Function GetValidValuesHTML(GroupKey, SelectedValue, bBlankOption)

	GetValidValuesHTML = ""

	if bBlankOption = true then 
		GetValidValuesHTML = GetValidValuesHTML & "<OPTION VALUE='' SELECTED>" 
	end if 
	
	Dim QSQL, Conn, Rs, Ids

	Conn = Session("ConnectionString")
	
	Set Rs = Server.CreateObject("ADODB.Recordset")

	If Not IsObject(Rs) Then
		Exit Function
	End If
	
    QSQL = "SELECT DISTINCT VV.Value FROM VALID_VALUE VV WHERE VV.Group_nm = '" & CStr(GroupKey) & "'"

	Rs.Open QSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Rs.EOF Then
		Exit Function
	End If

	Rs.MoveFirst
	    
	While Not Rs.EOF
		GetValidValuesHTML = GetValidValuesHTML & "<OPTION VALUE='" & Rs("Value") & "'" 
		
		if SelectedValue <> "" And SelectedValue = CStr(Rs("Value")) then
			GetValidValuesHTML = GetValidValuesHTML & " SELECTED>" 
		else
			GetValidValuesHTML = GetValidValuesHTML & ">"
		end if 

		GetValidValuesHTML = GetValidValuesHTML & Rs("Value")& "&nbsp"
		
		Rs.MoveNext
	Wend

    Rs.Close   
    Set Rs = Nothing
End Function
%>