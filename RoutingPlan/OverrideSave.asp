<!--#include file="..\lib\genericSQL.asp"-->
<%
on error resume next
	SQLUPDATE = Split(Request.Form("TxtUpdateData") , chr(130))
	SQLINSERT = Split(Request.Form("TxtInsertData") , chr(130))
	SQLDELETE = Split(Request.Form("TxtDeleteData") , chr(130))
		
	If 	Request.Form("TxtInsertData") <> "" Then
	For a = 0 to UBound(SQLINSERT)-1 step 1
		NewOVID = NextPkey("OUTPUT_MAPPING","OUTPUT_MAPPING_ID")
		INSERTSQL = BuildSQL(SQLINSERT(a),Chr(128), Chr(129), "INSERT", "OUTPUT_MAPPING", "OUTPUT_MAPPING_ID", NewOVID)		 
		Set RS = Conn.Execute(INSERTSQL)
	Next
	End If
	If 	Request.Form("TxtUpdateData") <> "" Then
	For X = 0 to UBound(SQLUPDATE)-1 step 1
		UpdateSQL = BuildSQL(SQLUPDATE(X),Chr(128), Chr(129), "UPDATE", "OUTPUT_MAPPING", "OUTPUT_MAPPING_ID", "")		 
		Set RS = Conn.Execute(UpdateSQL)
	Next
	
	End If
	If 	Request.Form("TxtDeleteData") <> "" Then
	For z = 0 to UBound(SQLDELETE)-1 step 1 
		DeleteSQL = BuildSQL(SQLDELETE(z),Chr(128), Chr(129), "DELETE", "OUTPUT_MAPPING","OUTPUT_MAPPING_ID" ,SQLDELETE(z))		 
		Set RS = Conn.Execute(DeleteSQL)
	Next
	End If
	
	strError = CheckADOErrors(Conn,"Override " & ACTION)
	If strError = "" Then
		Status = "Saved"
	Else
		Status = "Error"
	End If
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "CARRIER", "", 0, ""
		LogStatusGroupEnd
		%>
		'parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		'parent.frames('WORKAREA').SetDirty();
		'parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		'parent.frames('WORKAREA').UpdateStatus('Update successful.');
		'parent.frames('WORKAREA').ClearDirty();
		'parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If  %>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	'top.frames("WORKAREA").location.href = "RP_TabWindow.asp?OPID=<%= Request.QueryString("OPID") %>"
End Sub

-->
</SCRIPT>
</HEAD>
<BODY>

</BODY>
