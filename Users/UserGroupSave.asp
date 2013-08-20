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

	If ACTION = "INSERT" Then
		InsertSQL = ""
		InsertSQL = "INSERT INTO USER_GROUP (USER_ID, GROUP_ID) VALUES " & SQL_STRING
		Set RSUpdate = Conn.Execute(InsertSQL)
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = "DELETE FROM USER_GROUP WHERE " & SQL_STRING
		Set RSUpdate = Conn.Execute(DeleteSQL)
	End If 
	
	strError = CheckADOErrors(Conn,"User Group " & ACTION)
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "USER_GROUP", "", 0, ""
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
