<%
Function CheckSharedFrame(nFrmID, bLogToStatusRpt, bNewStatusRptGrp, nStatusRptMinCount, bLogToRefCountRpt, bNewRefCountRptGrp, nRefCountRptMinCount)

	CheckSharedFrame = 0

	Dim QSQL, Conn, Rs, Ids

	Conn = Session("ConnectionString")
	
	Set Rs = Server.CreateObject("ADODB.Recordset")

	If Not IsObject(Rs) Then
		Exit Function
	End If
	
    QSQL = "SELECT DISTINCT F.Frame_id, F.Name FName, FO.SEQUENCE, CF.CallFlow_id, CF.Name CFName FROM FRAME F, FRAME_ORDER FO, CALLFLOW CF WHERE F.Frame_id =" & CStr(nFrmID)  & " AND F.Frame_id = FO.Frame_id AND FO.CallFlow_id = CF.CallFlow_id"
	
	Rs.MaxRecords = Application("MAXIMUMSHAREDCOUNT")
	Rs.Open QSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Not Rs.EOF Then
		Rs.MoveLast
		CheckSharedFrame = Rs.RecordCount
	Else
		Exit Function
	End If
	

	Rs.MoveFirst

	If ((bLogToStatusRpt) And (CheckSharedFrame >= nStatusRptMinCount)) Then
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupBegin()
		End If
		    
		While Not Rs.EOF
			Ids = LogStatus(S_WARNING, "Frame " & Rs("FName") & " (" & Rs("Frame_id") & ") being referenced indirectly by Call Flow " & Rs("CFName") & " (" & Rs("CallFlow_id") & ").", "FRAME_ORDER", "Frame_id" , Rs("Frame_id"), "Seq #:" & Rs("SEQUENCE"))
			Rs.MoveNext
		Wend
   
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupEnd()
		End If
	End If
	
	If ((bLogToRefCountRpt) And (CheckSharedFrame >= nRefCountRptMinCount))Then
		Rs.MoveFirst
		
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupBegin()
		End If
		
		While Not Rs.EOF
			Ids = LogRefCount("Frame " & Rs("FName") & " (" & Rs("Frame_id") & ") being referenced indirectly by Call Flow " & Rs("CFName") & " (" & Rs("CallFlow_id") & ").", "FRAME_ORDER", "Frame_id" , Rs("Frame_id"), "Seq #:" & Rs("SEQUENCE"))
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