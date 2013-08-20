<%
Function CheckSharedOutputDef(nODID, bLogToStatusRpt, bNewStatusRptGrp, nStatusRptMinCount, bLogToRefCountRpt, bNewRefCountRptGrp, nRefCountRptMinCount)

	CheckSharedOutputDef = 0

	Dim QSQL, Conn, Rs, Ids

	Conn = Session("ConnectionString")
	
	Set Rs = Server.CreateObject("ADODB.Recordset")

	If Not IsObject(Rs) Then
		Exit Function
	End If
	
    QSQL = "SELECT DISTINCT OD.OUTPUTDEF_ID ODID, OD.Name ODName, OI.OUTPUT_ITEM_ID OIID, OI.SEQUENCE OISEQ, TSS.TRANSMISSION_SEQ_STEP_ID TSSID, RP.DESCRIPTION RPDESC, RP.ROUTING_PLAN_ID RPID FROM OUTPUT_DEFINITION OD, OUTPUT_ITEM OI, TRANSMISSION_SEQ_STEP TSS, ROUTING_PLAN RP WHERE OD.OUTPUTDEF_ID =" & CStr(nODID)  & " AND OD.OUTPUTDEF_ID = OI.OUTPUTDEF_ID AND OI.TRANSMISSION_SEQ_STEP_ID = TSS.TRANSMISSION_SEQ_STEP_ID AND TSS.ROUTING_PLAN_ID = RP.ROUTING_PLAN_ID "
	Rs.Open QSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Not Rs.EOF Then
		Rs.MoveLast
		CheckSharedOutputDef = Rs.RecordCount
	Else
		Exit Function
	End If	

	Rs.MoveFirst

	If ((bLogToStatusRpt) And (CheckSharedOutputDef >= nStatusRptMinCount)) Then
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupBegin()
		End If
		    
		While Not Rs.EOF
			Ids = LogStatus(S_WARNING, "Output Definition [" & Rs("ODName") & "] (" & Rs("ODID") & ") being referenced indirectly by Routing Plan [" & Rs("RPDESC") & "] (" & Rs("RPID") & "); Trans. Seq. Step (" & Rs("TSSID") & ").", "OUTPUT_ITEM", "OUTPUTDEF_ID" , Rs("OIID"), "Seq. # " & Rs("OISEQ"))
			Rs.MoveNext
		Wend
   
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupEnd()
		End If
	End If
	
	If ((bLogToRefCountRpt) And (CheckSharedOutputDef >= nRefCountRptMinCount))Then
		Rs.MoveFirst
		
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupBegin()
		End If
		
		While Not Rs.EOF
			Ids = LogRefCount("Output Definition [" & Rs("ODName") & "] (" & Rs("ODID") & ") being referenced indirectly by one or more Routing Plans.", "OUTPUT_ITEM", "OUTPUTDEF_ID" , Rs("OIID"), "Being referenced indirectly by Routing Plan [" & Rs("RPDESC") & "] (" & Rs("RPID") & "); Trans. Seq. Step (" & Rs("TSSID") & ").")
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