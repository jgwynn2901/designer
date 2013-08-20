<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
    '*************************************************
	' DMS: 3/21/00 
	' The DB design has been modified and there are 2 tables
	' CONTACT and AHS_CONTACT to update and get information from.
	'*************************************************
    

	On Error Resume Next

	dim NewAHSCOID, InsertAHSCSQL
	
	ACTION     = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")
	AHSID      = Request.Form("AHSID")		
	If ACTION  = "UPDATE" Then
		UpdateSQL    = ""
		UpdateSQL    = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "CONTACT", "CONTACT_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError     = CheckADOErrors(Conn,"CONTACT " & ACTION)
	Elseif ACTION = "INSERT" Then
		
		InsertSQL = ""
		NewCOID = NextPkey("CONTACT","CONTACT_ID")
		If NewCOID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "CONTACT", "CONTACT_ID", NewCOID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"CONTACT " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateCOID('" & NewCOID &  "');")	
		Else
			strError = "Unable to obtain next primary key for CONTACT table."
		End If	
	End If
	
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "CONTACT", "", 0, ""
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
