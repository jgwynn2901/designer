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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "OWNER", "OWNER_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Owner" & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewOID = CLng(NextPkey("OWNER","OWNER_ID"))
		If NewOID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "OWNER", "OWNER_ID", NewOID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Owner" & ACTION)
			If strError = "" Then 
			   Response.write("parent.frames('WORKAREA').UpdateOID('" & NewOID &  "');")
			end if
		Else
			strError = "Unable to obtain next primary key for OWNER table."
		End If			
	End if		
	
	Conn.Close
	
	If strError <> "" Then		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'></SPAN>Update unsuccessful, check status report.");
		parent.frames("WORKAREA").SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);		
<%	Else	%>
	    parent.frames("WORKAREA").UpdateStatus("Update successful.");
		parent.frames("WORKAREA").ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);				
<%	End If %>


</SCRIPT>
</HEAD>
<BODY>
</BODY>
