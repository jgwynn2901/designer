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
	
	strWarning = ""
	strOutputPageName = CStr(Request.Form("TxtName"))
		
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "OUTPUT_PAGE", "OUTPUT_PAGE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"OutputPage " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewOPID = CLng(NextPkey("OUTPUT_PAGE","OUTPUT_PAGE_ID"))
		If NewOPID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "OUTPUT_PAGE", "OUTPUT_PAGE_ID", NewOPID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"OutputPage " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateOPID('" & NewOPID &  "');")	
		Else
			strError = "Unable to obtain next primary key for OUTPUT_PAGE table."
		End If			
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "OUTPUT_PAGE", "OUTPUT_PAGE_ID", SQL_STRING)		 
		Set RSDelete = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"OutputPage " & ACTION)
	End If
	
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "OutputPage", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);

<%	ElseIf strWarning <> "" Then 
		LogStatusGroupBegin
		LogStatus S_WARNING, strWarning, "ATTRIBUTE", "", 0, ""
		LogStatusGroupEnd
%>		
parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Update successful with warnings, check status report.")	
parent.frames("WORKAREA").ClearDirty();
parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)		

<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		

		<%If ACTION	= "DELETE" Then%>
			parent.frames('WORKAREA').UpdateScreenOnDelete();
		<%End If%>

<%	End If

 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
