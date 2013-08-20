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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "OUTPUT_DEFINITION", "OUTPUTDEF_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Output Definition " & ACTION)		
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewODID = NextPkey("OUTPUT_DEFINITION", "OUTPUTDEF_ID")
		If NewODID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "OUTPUT_DEFINITION", "OUTPUTDEF_ID", NewODID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"Output Definition " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateODID(" & NewODID &  ");")	
		Else
			strError = "Unable to obtain next primary key for OUTPUT_DEFINITION table."
		End If				
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "OUTPUT_DEFINITION", "OUTPUTDEF_ID", SQL_STRING)		 
		Set RSUpdate = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Output Definition " & ACTION)
	ElseIf ACTION = "COPY" Then
		CopySQL = ""
		CopySQL = CopySQL & "{call Designer.CopyOutputDefinition(" & Request.Form("COPYODID") & ",{resultset 1, outOutputDefId, StatusMsg, StatusNum})}"
		Set RSCopy = Conn.Execute(CopySQL)
		NewODID = RSCopy("outOutputDefId")
		Response.write("parent.frames('WORKAREA').UpdateODID(" & RSCopy("outOutputDefId") &  ");")
		strError = CheckADOErrors(Conn,"Output Definition " & ACTION)
	End If
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "Output Definition", "", 0, ""
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
		<% If ACTION = "INSERT" OR ACTION = "COPY" Then %>
			parent.frames('WORKAREA').location.href = "OutputDefinitionDetails.asp?MODE=RW&ODID=<%= NewODID %>&SearchODID=<%= Request.Form("SearchODID")  %>&SearchNAME=<%= Request.Form("SearchNAME")  %>&SearchDESCRIPTION=<%= Request.Form("SearchDESCRIPTION")  %>&SearchType=<%= Request.Form("SearchType") %>"
	<% End If %>
<%	End If  %>
</SCRIPT>
</HEAD>
<BODY>
</BODY>
