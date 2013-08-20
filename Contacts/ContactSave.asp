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
	DeleteSQL = ""
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "CONTACT", "CONTACT_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Contact " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewCID = CLng(NextPkey("CONTACT","CONTACT_ID"))
		If NewCID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "CONTACT", "CONTACT_ID", NewCID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Contact " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateCID('" & NewCID &  "');")
		Else
			strError = "Unable to obtain next primary key for CONTACT table."
		End If			
			
	ElseIf len(Request.Querystring("DELETE")) <> 0 Then
		DeleteSQL = "Delete from CONTACT Where CONTACT_ID=" & Request.Querystring("DELETE")
		Set RSUpdate = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Contact " & ACTION)
	End If 
	Conn.Close
	if len(DeleteSQL) <> 0 then
		'If strError <> "" Then
		'	response.Write "<script language=vbscript>Msgbox """ & strError & """, 16" & VBcrlf & "</script>"
		'else
			response.redirect "ContactDetailsData.asp?AHSID=" & Request.Querystring("AHSID")
		'end if
	else		
		If strError <> "" Then
			LogStatusGroupBegin
			LogStatus S_ERROR, strError, "CONTACT", "", 0, ""
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
	end if
 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
