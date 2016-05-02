<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<%
	Response.Expires = 0
	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")		
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "MIGRATION_JOB", "JOB_ID", "")
		Set RSinsert = Conn.Execute(UpdateSQL)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewJID = NextPkey("MIGRATION_JOB","JOB_ID")
		InsertSQL = BuildSQL(SQL_STRING, Chr(124), Chr(126), "INSERT", "MIGRATION_JOB", "JOB_ID", NewJID)
		Set RSUpdate = Conn.Execute(InsertSQL)

		If Request.Form("AllRoutingRelatedItems") = "on" Then
			NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 9,0)"
			Set RSUpdate = Conn.Execute(InsertSQL)
		Else
			ArrayRP = Split(Request.Form("HiddenALlRoutingRelatedItems"), ",")
			For r = 0 to UBound(ArrayRP) step 1
				NewDID = ""
				NewDID = NextPKey("MIGRATION_DETAIL","JOB_DETAIL_ID")
				InsertSQL = ""
				InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
				InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 9," & ArrayRP(r) & ")"
				Set RSUpdate = Conn.Execute(InsertSQL)
			next
		End If
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)		
		
		If Request.form("ALLDEFINITIONS") = "on" Then
			NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 8,0)"
			Set RSUpdate = Conn.Execute(InsertSQL)
		Else		
			ArrayOD = split(Request.Form("HiddenODlist"), ",")
			for i = 0 to Ubound(ArrayOD) step 1
				NewDID = ""
				NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
				InsertSQL = ""
				InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
				InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 8," & ArrayOD(i) & ")"
				Set RSUpdate = Conn.Execute(InsertSQL)
			next
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("RULES") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 1," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("RULES_LOOKUPS") = "on" Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 10," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("CLAIMNUMBER") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 2," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)		

		If Request.form("ATTRIBUTES") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 11," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("ASSIGNMENT") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 3," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("ACCOUNTCALLFLOW") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 12," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("ROUTINGADDRESS") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 4," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("INFORMATION") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 6," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If

strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("ESCALATION_RULES") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 5," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("FEE") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL  (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 0," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("COMMON") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 7," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("OutputOverFlow") = "on"  Then
		NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 14," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		
strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("EDIROUTING") = "on"  Then
			NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = ""
			InsertSQL = InsertSQL & "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 13," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
	
		strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("VendorDefs") = "on"  Then
			NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 15," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)

		If Request.form("FraudDefs") = "on"  Then
			NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 16," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
		
		If Request.form("Subrogation") = "on"  Then
			NewDID = NextPkey("MIGRATION_DETAIL","JOB_DETAIL_ID")
			InsertSQL = "INSERT INTO MIGRATION_DETAIL (JOB_ID, JOB_DETAIL_ID, SUBSET_ID, ID_TO_MOVE) VALUES ("
			InsertSQL = InsertSQL & NewJID & "," & NewDID & ", 17," & "null" & ")"
			Set RSUpdate = Conn.Execute(InsertSQL)
		End If
		strError = strError & CheckADOErrors(Conn,"Migration Job" & ACTION)
	End If
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "CARRIER", "", 0, ""
		LogStatusGroupEnd
		%>
<HTML>
<HEAD>
</HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<BODY  BGCOLOR="<%=BODYBGCOLOR%>">
<SPAN CLASS="LABEL"> Migration save was unsuccessful:</SPAN><br>
<SPAN CLASS="LABEL"><%=strError%></SPAN><br>
<BUTTON NAME=BtnBack CLASS=STDBUTTON onclick="window.history.back();" ><U>B</U>ack</BUTTON>
</BODY>

<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd 
		PARAMS = "?START_TIME=" & Request.Form("START_TIME")
		PARAMS = PARAMS & "&START_DATE="&Request.Form("START_DATE")
		PARAMS = PARAMS & "&CONNECT_STRING=" & Session("ConnectionString") 
		PARAMS = PARAMS & "&JOB_ID=" & NewJID
		PARAMS = PARAMS & "&SCHEDULER_APP_PATH=" & Application("SchedulerAppPath")
		PARAMS = PARAMS & "&MIGRATION_SERVER=" & Application("MigrationServer")
		PARAMS = PARAMS & "&MIGRATION_APP_PATH=" & Application("MigrationAppPath")
		Response.Redirect "MigrationConfirm.asp" & PARAMS
	End If  %>

