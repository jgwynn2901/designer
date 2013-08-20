<%
Function CheckSharedCallFlow(nCFID, bLogToStatusRpt, bNewStatusRptGrp, nStatusRptMinCount, bLogToRefCountRpt, bNewRefCountRptGrp, nRefCountRptMinCount)


	CheckSharedCallFlow = 0

	Dim QSQL, Conn, Rs, Ids

	Conn = Session("ConnectionString")
	
	Set Rs = Server.CreateObject("ADODB.Recordset")

	If Not IsObject(Rs) Then
		Exit Function
	End If
	
    QSQL = "SELECT DISTINCT CF.CallFLow_id, CF.Name CFName, AHS.Name AHSName, AHS.ACCNT_HRCY_STEP_ID AHSID, ACF.ACCOUNTCALLFLOW_ID, ACF.LOB_CD FROM CALLFLOW CF, ACCOUNT_CALLFLOW ACF, ACCOUNT_HIERARCHY_STEP AHS WHERE CF.CallFlow_id =" & CStr(nCFID)  & " AND CF.CallFlow_id = ACF.CallFlow_id AND ACF.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID"

	Rs.MaxRecords = Application("MAXIMUMSHAREDCOUNT")
	Rs.Open QSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Not Rs.EOF Then
		Rs.MoveLast
		CheckSharedCallFlow = Rs.RecordCount
	Else
		Exit Function
	End If
	

	Rs.MoveFirst

	If ((bLogToStatusRpt) And (CheckSharedCallFlow >= nStatusRptMinCount)) Then
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupBegin()
		End If
		    
		While Not Rs.EOF
			Ids = LogStatus(S_WARNING, "Call Flow " & Rs("CFName") & " (" & Rs("CallFLow_id") & ") being referenced indirectly by Account Hierarchy Step " & Rs("AHSName") & " (" & Rs("AHSID") & ").", "ACCOUNT_CALLFLOW", "CallFLow_id" , Rs("ACCOUNTCALLFLOW_ID"), "LOB code:" & Rs("LOB_CD"))
			Rs.MoveNext
		Wend
   
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupEnd()
		End If
	End If
	
	If ((bLogToRefCountRpt) And (CheckSharedCallFlow >= nRefCountRptMinCount))Then
		Rs.MoveFirst
		
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupBegin()
		End If
		
		While Not Rs.EOF
			Ids = LogRefCount("Call Flow " & Rs("CFName") & " (" & Rs("CallFLow_id") & ") being referenced indirectly by Account Hierarchy Step " & Rs("AHSName") & " (" & Rs("AHSID") & ").", "ACCOUNT_CALLFLOW", "CallFLow_id" , Rs("ACCOUNTCALLFLOW_ID"), "LOB code:" & Rs("LOB_CD"))
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