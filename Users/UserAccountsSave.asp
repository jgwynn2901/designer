<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	strError = ""
	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")	
	If ACTION = "INSERT" Then
		InsertSQL = SQL_STRING
		Set RSUpdate = Conn.Execute(InsertSQL)
		if CStr(RSUpdate("StatusNum")) <> "0" then
			strError = RSUpdate("StatusMsg")
		End If
	End If		
	If ACTION = "DELETE" Then
		DeleteSQL = "DELETE FROM ACCOUNT_USER WHERE " & SQL_STRING
		Set RSUpdate = Conn.Execute(DeleteSQL)
	End If 
	
	If strError <> "" Then	strError = strError & VBCRLF 
	strError = strError & CheckADOErrors(Conn,"Account User " & ACTION)
		
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "ACCOUNT_USER", "", 0, ""
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames("WORKAREA").SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);		
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("Update successful.");
		parent.frames("WORKAREA").ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);
<%	End If
 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
