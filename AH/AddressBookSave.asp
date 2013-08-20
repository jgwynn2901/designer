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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "ADDRESS_BOOK_ENTRY", "ADDRESS_BOOK_ENTRY_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"ADDRESS_BOOK_ENTRY " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewABID = NextPkey("ADDRESS_BOOK_ENTRY", "ADDRESS_BOOK_ENTRY_ID")
		If NewABID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "ADDRESS_BOOK_ENTRY", "ADDRESS_BOOK_ENTRY_ID", NewABID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"ADDRESS_BOOK_ENTRY " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateABID(" & NewABID &  ");")	
		Else
			strError = "Unable to obtain next primary key for ADDRESS_BOOK_ENTRY table."
		End If			
					

	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "ADDRESS_BOOK_ENTRY", "ADDRESS_BOOK_ENTRY_ID", SQL_STRING)		 
		Set RSUpdate = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"ADDRESS_BOOK_ENTRY " & ACTION)
	End If
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "ADDRESS_BOOK_ENTRY", "", 0, ""
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
		<%If ACTION	= "DELETE" Then%>
			parent.frames('WORKAREA').UpdateScreenOnDelete();
		<%End If%>
<%	End If  %>
</SCRIPT>
</HEAD>
<BODY>
</BODY>
