<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")		
	If ACTION = "UPDATE" Then
		UpdateSQL = "UPDATE ACCOUNT_HIERARCHY_STEP set ADDRESS_1='" & request("TxtStreet") & "'," & _
			"CITY='" & request("City") & "'," & _ 
			"STATE='" & request("State") & "'," & _   
			"ZIP='" & request("ZIP") & "' " & _
			"WHERE ACCNT_HRCY_STEP_ID = " & request("AHS_ID")
		Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"ACCOUNT_HIERARCHY_STEP " & ACTION)
	elseif ACTION = "INSERT" Then
		strError = doInsert
	ElseIf ACTION = "DELETE" Then
		strError = doDelete( request("AHSID"), request("AHS_POLID") )
	End If 
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "ACCOUNT_HIERARCHY_STEP", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus('Update unsuccessful, check status report.');
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If  

function doDelete( cAHSID, cAHS_POLID )
dim cSQL, cError

On Error Resume Next
Conn.BeginTrans
cSQL = "DELETE FROM AHS_POLICY WHERE AHS_POLICY_ID = " & cAHS_POLID
Conn.execute cSQL
if err.number = 0 then
	cSQL = "DELETE FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID = " & cAHSID
	Conn.execute cSQL
	if err.number <> 0 then
		cError = CheckADOErrors(Conn,"AHS_POLICY" & ACTION)
	end if
else
	cError = CheckADOErrors(Conn,"ACCOUNT_HIERARCHY_STEP" & ACTION)
end if		
If cError <> "" Then
	Conn.RollbackTrans
else
	Conn.CommitTrans
End If
doDelete = cError
end function

function doInsert
dim cSQL, cError, cNew_ID, cNew_AHSID

On Error Resume Next
Conn.BeginTrans
cNew_AHSID = CLng(NextPkey("ACCOUNT_HIERARCHY_STEP", "ACCNT_HRCY_STEP_ID"))
If cNew_AHSID <= 0 Then
	cError = "Unable to obtain next primary key for ACCOUNT_HIERARCHY_STEP table."
else
	cSQL = "INSERT INTO ACCOUNT_HIERARCHY_STEP " & _
			"(ACCNT_HRCY_STEP_ID," & _
			"ADDRESS_1," & _
			"CITY," & _
			"STATE," & _
			"ZIP," & _
			"AUTO_ESCALATE," & _
			"ACTIVE_STATUS," & _
			"NODE_TYPE_ID) " & _
	"VALUES ( " & cNew_AHSID & ",'" & _
			request("TxtStreet") & "','" & _
			request("City") & "','" & _ 
			request("State") & "','" & _   
			request("ZIP") & "'," & _
			"'N'," & _
			"'ACTIVE'," & _
			"2 )"
	Conn.execute cSQL
	if err.number = 0 then
		cNew_ID = CLng(NextPkey("AHS_POLICY", "AHS_POLICY_ID"))
		if cNew_ID > 0 then
			cSQL = "INSERT INTO AHS_POLICY " & _
					"(AHS_POLICY_ID," & _
					"ACCNT_HRCY_STEP_ID," & _
					"POLICY_ID," & _
					"LOB_CD) " & _
					"VALUES ( " & cNew_ID & "," & _
					cNew_AHSID & ",'" & _
					request("POLICY_ID") & "','" & _ 
					request("LOB") & "')"
			Conn.execute cSQL
			cError = CheckADOErrors(Conn,"Ahs_Policy " & ACTION)
		else			
			cError = "Unable to obtain next primary key for AHS_POLICY table."
		End If
	else
		cError = CheckADOErrors(Conn,"ACCOUNT_HIERARCHY_STEP" & ACTION)
	end if
end if		
If cError <> "" Then
	Conn.RollbackTrans
else
	Conn.CommitTrans
	Response.write("parent.frames('WORKAREA').UpdateAHSID('" & cNew_ID &  "');")	
End If
doInsert = cError
end function
%>
</SCRIPT>
</HEAD>
<BODY>
</BODY>
