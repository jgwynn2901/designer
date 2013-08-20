<%
Function CheckSharedLUType(nLUTypeID, bLogToStatusRpt, bNewStatusRptGrp, nStatusRptMinCount, bLogToRefCountRpt, bNewRefCountRptGrp, nRefCountRptMinCount)


	CheckSharedLUType = 0

	Dim QSQL, Conn, Rs, Ids

	Conn = Session("ConnectionString")
	
	Set Rs = Server.CreateObject("ADODB.Recordset")

	If Not IsObject(Rs) Then
		Exit Function
	End If
	
    QSQL = " SELECT DISTINCT L.LU_TYPE_ID, L.Name LUName, A.Name AName, A.Attribute_id " &_
           "   FROM ATTRIBUTE A, ATTR_INSTANCE AI, LU_TYPE L , ATTRIBUTE_OVERRIDE AO " &_
           "  WHERE L.LU_TYPE_ID =" & CStr(nLUTypeID)  &_
           "    AND AI.LU_TYPE_ID = L.LU_TYPE_ID " &_
           "    AND A.ATTRIBUTE_ID = AI.ATTRIBUTE_ID " &_
           "    AND AI.ATTR_INSTANCE_ID = AO.ATTR_INSTANCE_ID(+) "
    
	Rs.MaxRecords = Application("MAXIMUMSHAREDCOUNT")
	Rs.Open QSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Not Rs.EOF Then
		Rs.MoveLast
		CheckSharedLUType = Rs.RecordCount
	Else
		Exit Function
	End If	

	Rs.MoveFirst

	If ((bLogToStatusRpt) And (CheckSharedLUType >= nStatusRptMinCount)) Then
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupBegin()
		End If
		    
		While Not Rs.EOF
			Ids = LogStatus(S_WARNING, "Look Up Type " & Rs("LUName") & " (" & Rs("LU_TYPE_ID") & ") being referenced indirectly by Attribute " & Rs("AName") & " (" & Rs("Attribute_id") & ").", "ATTRIBUTE", "LU_TYPE_ID" , Rs("Attribute_id"), Rs("AName"))
			Rs.MoveNext
		Wend
   
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupEnd()
		End If
	End If
	
	If ((bLogToRefCountRpt) And (CheckSharedLUType >= nRefCountRptMinCount))Then
		Rs.MoveFirst
		
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupBegin()
		End If
		
		While Not Rs.EOF
			Ids = LogRefCount("Look Up Type " & Rs("LUName") & " (" & Rs("LU_TYPE_ID") & ") being referenced indirectly by one or more Attributes", "ATTRIBUTE", "LU_TYPE_ID" , Rs("Attribute_id"), Rs("AName"))
			Rs.MoveNext
		Wend
   
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupEnd()
		End If
    End If

    Rs.Close   
    Set Rs = Nothing
End Function
%>