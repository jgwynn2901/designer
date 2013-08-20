<%
Function CheckSharedAgent(nAgentID, bLogToStatusRpt, bNewStatusRptGrp, nStatusRptMinCount, bLogToRefCountRpt, bNewRefCountRptGrp, nRefCountRptMinCount)



	CheckSharedAgent = 0

	Dim QSQL, Cn, CPw1, ParamInA1, Conn, Rs, Ids

	Conn = Session("ConnectionString")

    Set Cn = Server.CreateObject("ADODB.Connection")
    Cn.ConnectionString = Conn
    Cn.CursorLocation = adUseClient
    Cn.Open
    
    QSQL = "{call Designer_2.GetAgentUsedBy(" & nAgentID & "," & Application("MAXIMUMSHAREDCOUNT") & ",{resultset 2000, AID, AIDName, UsedInTable, UsedInField, UsedInRowID, UsedInRowDesc, RecordCount})}"
            
    Set CPw1 = Server.CreateObject("ADODB.Command")
    CPw1.ActiveConnection = Cn
    CPw1.CommandText = QSQL
    CPw1.CommandType = adCmdText

    Set Rs = CPw1.Execute
	
	If Not IsObject(Rs) Then
		Exit Function
	End If

	If Not Rs.EOF Then
		CheckSharedAgent = CLng(Rs(6))
	Else
		Exit Function
	End If

	If (bLogToStatusRpt) And (CheckSharedAgent >= nStatusRptMinCount) Then
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupBegin()
		End If
    
		While Not Rs.EOF
			Ids = LogStatus(S_WARNING, "Agent (" & Rs(0) & ") being referenced by more than 1 item.", Rs(2), Rs(3), Rs(4), Rs(5))
			Rs.MoveNext
		Wend
		If (bNewStatusRptGrp) Then
			Call LogStatusGroupEnd()
		End If
	End If

	If (bLogToRefCountRpt) And (CheckSharedAgent >= nRefCountRptMinCount) Then
		Rs.MoveFirst
		
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupBegin()
		End If
		
		While Not Rs.EOF
			Ids = LogRefCount("Agent (" & Rs(0) & ") being referenced by more than 1 item.", Rs(2), Rs(3), Rs(4), Rs(5))
			Rs.MoveNext
		Wend
   
		If (bNewRefCountRptGrp) Then
			Call LogRefCountGroupEnd()
		End If
    End If


    Rs.Close
   
    Set Rs = Nothing
	Set Cn = Nothing
	Set CPw1 = Nothing


End Function
%>