<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	dim lIsDuplicate
	dim nST, cLOB, nAHS, nAVID
	dim nOldST, cOldLOB

	lIsDuplicate = false
	On Error Resume Next
	nST = Request("ST")
	cLOB = Request("LOB")
	nAHS = Request("AHS_ID")
	nAVID = Request("AVID")
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")		
	If ACTION = "UPDATE" Then
		'	get current LOB and ST
		cSQL = "SELECT * FROM ACCOUNT_VENDOR " & _
				"WHERE ACCOUNT_VENDOR_ID = " & nAVID
		Set oRS = Conn.Execute(cSQL)
		nOldST = oRS("SERVICE_TYPE_ID")
		cOldLOB = oRS("LOB")
		oRS.close
		cSQL = "SELECT * FROM ACCOUNT_VENDOR " & _
				"WHERE ACCNT_HRCY_STEP_ID = " & nAHS & _
				" AND SERVICE_TYPE_ID = " & nST & _
				" AND LOB = '" & cLOB & "' AND ACCOUNT_VENDOR_ID <> " & nAVID
		Set oRS = Conn.Execute(cSQL)
		if not oRS.eof then
			lIsDuplicate = true
		else
			UpdateSQL = "UPDATE ACCOUNT_VENDOR SET LOB='" & cLOB & "', SERVICE_TYPE_ID=" & nST & _
					" Where ACCNT_HRCY_STEP_ID = " & nAHS & " AND SERVICE_TYPE_ID=" & nOldST & " AND LOB='" & cOldLOB & "'"
			Conn.Execute(UpdateSQL)
			strError = CheckADOErrors(Conn,"Account Vendors " & ACTION)
		end if
		oRS.close
	Elseif ACTION = "INSERT" Then
		cSQL = "SELECT * FROM ACCOUNT_VENDOR " & _
				"WHERE ACCNT_HRCY_STEP_ID = " & nAHS & _
				" AND SERVICE_TYPE_ID = " & nST & _
				" AND LOB = '" & cLOB & "'"
		Set oRS = Conn.Execute(cSQL)
		if not oRS.eof then
			lIsDuplicate = true
		else
			NewAVID = NextPkey("ACCOUNT_VENDOR", "ACCOUNT_VENDOR_ID")
			If NewAVID > 0 Then 
				InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "ACCOUNT_VENDOR", "ACCOUNT_VENDOR_ID", NewAVID)		 
				Set RSUpdate = Conn.Execute(InsertSQL)
				strError = CheckADOErrors(Conn,"Account Vendors " & ACTION)
				If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateAVID(" & NewAVID &  ");")	
			Else
				strError = "Unable to obtain next primary key for ACCOUNT_VENDOR table."
			End If			
		end if
		oRS.close
		set oRS = nothing
	Elseif ACTION = "DELETE" Then
		DeleteSQL = BuildSQL("", "", "", "DELETE", "ACCOUNT_VENDOR", "ACCOUNT_VENDOR_ID", SQL_STRING)		 
		Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"ACCOUNT_VENDOR " & ACTION)
	End If
	Conn.Close
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "Account Vendors", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Elseif lIsDuplicate then %>
		alert( "The combination of LOB/Service Type already exists!." );
<%	else	
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If%> 
</SCRIPT>
</HEAD>
<BODY>
</BODY>
