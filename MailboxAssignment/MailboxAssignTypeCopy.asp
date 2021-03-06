<%
'***************************************************************
'calls stored procedure to copy Mailbox Assignment Types
'
'$History: MailboxAssignTypeCopy.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:47p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'***************************************************************
%>
<%	Response.Expires = 0
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT LANGUAGE="VBScript">

<%	
	bErrors = True
		
	strMATID		= CStr(Request.QueryString("MATID"))
	strAHSID		= CStr(Request.QueryString("TxtAHSID"))
	
	strExecute = "{call Designer.SP_COPY_MAILBOXASSIGNMENTTYPE(" &_
					"'" & strMATID & "','" & strAHSID & "', {resultset  1, outResult})}"
				
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
	Set RS = Server.CreateObject("ADODB.Recordset")
	rs.Open strExecute,Conn ,adOpenStatic,adLockReadOnly, adCmdText
				
	if rs("outResult")="0" then
		bErrors = true
	else
		strMATID = rs("outResult") %>
		parent.frames("WORKAREA").UpdateMATID(<%=strMATID%>)
		parent.frames("WORKAREA").Refresh
<%		bErrors = false
	end if
			
	
	strError = CheckADOErrors(Conn,"Mailbox Assignment Type COPY" )
	RS.Close
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing

	If bErrors = True Or strError <> "" Then 
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "MAILBOX_ASSIGNMENT_TYPE", "", 0, ""
		LogStatusGroupEnd
%>		
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.")	
		parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)
		parent.frames("WORKAREA").SetDirty
<%	Else  
		LogStatusGroupBegin
		LogStatusGroupEnd
%>
		parent.frames("WORKAREA").UpdateStatus("Update successful.")			
		parent.frames("WORKAREA").SetStatusInfoAvailableFlag(false)
		parent.frames("WORKAREA").ClearDirty
<%	End If 
			
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>


