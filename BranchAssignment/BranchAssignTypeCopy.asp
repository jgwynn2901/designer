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
		
	strBATID		= CStr(Request.QueryString("BATID"))
	strAHSID		= CStr(Request.QueryString("TxtAHSID"))
	
	strExecute = "{call Designer.SP_COPY_BRANCHASSIGNMENTTYPE(" &_
					"'" & strBATID & "','" & strAHSID & "', {resultset  1, outResult})}"
				
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
	Set RS = Server.CreateObject("ADODB.Recordset")
	rs.Open strExecute,Conn ,adOpenStatic,adLockReadOnly, adCmdText
				
	if rs("outResult")="0" then
		bErrors = true
	else
		strBATID = rs("outResult") %>
		parent.frames("WORKAREA").UpdateBATID(<%=strBATID%>)
		parent.frames("WORKAREA").Refresh
<%		bErrors = false
	end if
			
	
	strError = CheckADOErrors(Conn,"Branch Assignment Type COPY" )
	RS.Close
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing

	If bErrors = True Or strError <> "" Then 
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "BRANCHASSIGNMENTTYPE", "", 0, ""
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


