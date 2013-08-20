<%
Function GetControlDataHTML(TableName, KeyCol, DisplayCol, SelectedValue, bBlankOption)

	GetControlDataHTML = ""

	if bBlankOption = true then 
		GetControlDataHTML = GetControlDataHTML & "<OPTION VALUE='' SELECTED>" 
	end if 
	
	Dim QSQL, Conn, Rs, Ids

	Conn = Session("ConnectionString")
	
	Set Rs = Server.CreateObject("ADODB.Recordset")

	If Not IsObject(Rs) Then
		Exit Function
	End If
	
    QSQL = "SELECT " & DisplayCol & " ," & KeyCol & " FROM " & TableName 

	Rs.Open QSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Rs.EOF Then
		Exit Function
	End If

	Rs.MoveFirst
	    
	While Not Rs.EOF
		GetControlDataHTML = GetControlDataHTML & "<OPTION VALUE='" & Rs(KeyCol) & "'" 
		
		if SelectedValue <> "" And Trim(SelectedValue) = Trim(CStr(Rs(DisplayCol))) then
			GetControlDataHTML = GetControlDataHTML & " SELECTED>" 
		else
			GetControlDataHTML = GetControlDataHTML & ">"
		end if 

		GetControlDataHTML = GetControlDataHTML & Rs(DisplayCol)& "&nbsp"
		
		Rs.MoveNext
	Wend

    Rs.Close   
    Set Rs = Nothing
End Function
%>