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
	strAttributeName = CStr(Request.Form("TxtName"))

	CheckSQL = "{call Designer_2.CheckAttrName(" &_	
				"'" & strAttributeName & "', {resultset  2, cStatusMsg, nStatusCode)}"
	Set RSCheck = Conn.Execute(CheckSQL)			
	if CStr(RSCheck("nStatusCode")) <> "0" then
		strWarning = RSCheck("cStatusMsg")
		strWarning = strWarning & VBCRLF & CheckADOErrors(Conn,"Attribute CheckAttrName")
	End If
	RSCheck.Close
			
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "ATTRIBUTE", "ATTRIBUTE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Attribute " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewAID = CLng(NextPkey("ATTRIBUTE","ATTRIBUTE_ID"))
		If NewAID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "ATTRIBUTE", "ATTRIBUTE_ID", NewAID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Attribute " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateAID('" & NewAID &  "');")	
		Else
			strError = "Unable to obtain next primary key for ATTRIBUTE table."
		End If			
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "ATTRIBUTE", "ATTRIBUTE_ID", SQL_STRING)		 
		Set RSDelete = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Attribute " & ACTION)
	End If

	Conn.Close
	
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "ATTRIBUTE", "", 0, ""
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
