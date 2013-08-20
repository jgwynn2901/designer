<%
'***************************************************************
'generates insert or update query for Mailbox Assignment Types 
'
'$History: MailboxAssignTypeSave.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:47p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'***************************************************************
%>
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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "MAILBOX_ASSIGNMENT_TYPE", "MAILBOX_ASSIGNMENT_TYPE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Mailbox Assignment Type " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewMATID = CLng(NextPkey("MAILBOX_ASSIGNMENT_TYPE","MAILBOX_ASSIGNMENT_TYPE_ID"))

		If NewMATID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "MAILBOX_ASSIGNMENT_TYPE", "MAILBOX_ASSIGNMENT_TYPE_ID", NewMATID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Mailbox Assignment Type " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateMATID('" & NewMATID &  "');")	
		Else
			strError = "Unable to obtain next primary key for MAILBOXASSIGNMENTTYPE table."
		End If			
		
	End If
	Conn.Close
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "MAILBOX_ASSIGNMENT_TYPE", "", 0, ""
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
