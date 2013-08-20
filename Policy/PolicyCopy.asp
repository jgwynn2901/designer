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

	strPID		= CStr(Request.QueryString("PID"))
	strAHSID		= CStr(Request.QueryString("TxtAHSID"))
	
	strExecute = "{call Designer.CopyPolicy(" &_
					"'" & strPID & "','" & strAHSID & "', {resultset  1, outPolicyID})}"
				
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
	Set RS = Server.CreateObject("ADODB.Recordset")
	rs.Open strExecute,Conn ,adOpenStatic,adLockReadOnly, adCmdText
				
	if rs("outPolicyID")="0" then
		bErrors = true
	else
		strPID = rs("outPolicyID") %>
		parent.frames("WORKAREA").UpdatePID(<%=strPID%>)
		parent.frames("WORKAREA").Refresh
<%		bErrors = false
	end if
			
	
	strError = CheckADOErrors(Conn,"Policy COPY" )
	RS.Close
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing

	If bErrors = True Or strError <> "" Then 
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "POLICY", "", 0, ""
		LogStatusGroupEnd
%>		
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.")	
		parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)
<%	Else  
		LogStatusGroupBegin
		LogStatusGroupEnd
%>
		parent.frames("WORKAREA").UpdateStatus("Update successful.")			
		parent.frames("WORKAREA").SetStatusInfoAvailableFlag(false)
<%	End If 
			
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>


