<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
dim cSQL, cAction, cError, oConn, nRecs

On Error Resume Next
cAction = Request.Form("txtAction")
cSQL = Request.Form("txtSaveData")	
	
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
oConn.Execute cSQL, nRecs, adCmdText
if err.number <> 0 then
	if err.number = -2147217900 then
		' key is not unique
		cError = "Client Code/Policy ID combination must be unique."
	else
		cError = CheckADOErrors(oConn, "iNetPOLICY " & cAction)
	end if
end if	
oConn.Close
set oConn = nothing
If cError <> "" Then
	LogStatusGroupBegin
	LogStatus S_ERROR, cError, "iNetPOLICY", "", 0, ""
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
