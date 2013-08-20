<%
'***************************************************************
'generates insert or update query for Mailbox Assignment Rules 
'
'$History: MailboxAssignRuleSave.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:46p
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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "MAILBOX_ASSIGNMENT_RULE", "MAILBOX_ASSIGNMENT_RULE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Mailbox Assignment Rule " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewMARID = CLng(NextPkey("MAILBOX_ASSIGNMENT_RULE","MAILBOX_ASSIGNMENT_RULE_ID"))
		If NewMARID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "MAILBOX_ASSIGNMENT_RULE", "MAILBOX_ASSIGNMENT_RULE_ID", NewMARID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Mailbox Assignment Rule " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateMARID('" & NewMARID &  "');")
		Else
			strError = "Unable to obtain next primary key for MAILBOX_ASSIGNMENT_RULE table."
		End If			
			
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "MAILBOX_ASSIGNMENT_RULE", "MAILBOX_ASSIGNMENT_RULE_ID", SQL_STRING)		 
		Set RSUpdate = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Mailbox Assignment Rule " & ACTION)
		
		'Trilok - Audit Delete Changes.
		DeleteSQL = BuildDeleteAuditSQL("MAILBOX_ASSIGNMENT_RULE_AUDIT","MAILBOX_ASSIGNMENT_RULE_ID",SQL_STRING)
	  Conn.Execute(DeleteSQL)
	  
	End If 
	Conn.Close
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "MAILBOX_ASSIGNMENT_RULE", "", 0, ""
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames("WORKAREA").SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);		
<%		
	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("Update successful.");
		parent.frames("WORKAREA").ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);				
<%		
	End If

 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
