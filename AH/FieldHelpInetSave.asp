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
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "CALL_HELP", "HELP_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"CALL_HELP " & ACTION)
Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewHPID = NextPkey("CALL_HELP","HELP_ID")
		
		If NewHPID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "CALL_HELP", "HELP_ID", NewHPID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"CALL_HELP " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateHPID('" & NewHPID &  "');")	
		Else
			strError = "Unable to obtain next primary key for CALL_HELP table."
		End If			
		
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "CALL_HELP", "HELP_ID", SQL_STRING)		 
		Set RSDelete = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Attribute " & ACTION)
	End If
	
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "CALL_HELP", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If  %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
