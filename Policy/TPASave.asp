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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "THIRD_PARTY_ADMINISTRATOR", "TPA_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"THIRD_PARTY_ADMINISTRATOR " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewCID = NextPkey("THIRD_PARTY_ADMINISTRATOR","TPA_ID")
		If NewCID > 0 Then 			
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "THIRD_PARTY_ADMINISTRATOR", "TPA_ID", NewCID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"THIRD_PARTY_ADMINISTRATOR " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateCID('" & NewCID &  "');")	
		Else
			strError = "Unable to obtain next primary key for THIRD_PARTY_ADMINISTRATOR table."
		End If				
			
	End If
	
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "TPA", "", 0, ""
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
