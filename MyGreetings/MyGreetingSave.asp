<%
'***************************************************************
'generates insert or update query for Mailboxes.
'
'$History: MyGreetingSave.asp $ 
'* 
'* *****************  Version 4  *****************
'* User: Jenny.cheung Date: 7/09/08    Time: 3:43p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* REMOVED STOP
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:35p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:33p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:28p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* took out stop
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:25p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:14p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:09p
'* Created in $/FNS_DESIGNER/Source/Designer/Greeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 4/21/08    Time: 9:23a
'* Created in $/FNS_DESIGNER/Source/Designer
'* created for Sedgwick.  Just want to save my work for now
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:46p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Mailbox
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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "GREETINGS", "GREETINGS_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Greetings" & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewGreetingID = CLng(NextPkey("GREETINGS","GREETINGS_ID"))
		If NewGreetingID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "GREETINGS", "GREETINGS_ID", NewGreetingID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"GREETINGS" & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateGreetingID('" & NewGreetingID &  "');")	
		Else
			strError = "Unable to obtain next primary key for Greeting table."
		End If			

	End If

	Conn.Close
	
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "GREETINGS", "", 0, ""
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
<%	End If

 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
