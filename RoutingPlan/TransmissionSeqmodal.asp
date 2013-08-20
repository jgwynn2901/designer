<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
If HasViewPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "Override_Layout_Bottom.asp"
End If
If HasModifyPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then MODE = "RO"
%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Transmission Sequence Step</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	
End Sub
Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub BtnCancel_OnClick
MsgRet = "1"


If Parent.frames("SeqIFrame").document.frames("TOP").document.body.getAttribute("CanDocUnloadNowInf") = "NO" Then
	MsgRet = msgbox ("Data has changed, leave page without saving?", 1, "FNSNetDEsigner")
End If
	if MsgRet="1" Then
		window.returnvalue = "CANCEL"
		window.close
	End If
End Sub

Sub BtnSave_onclick
Set DataPage = Parent.frames("SeqIFrame").document.frames("TOP").document.all
ErrMSg = ""

if DataPage.SEQUENCE.value = "" OR Not IsNumeric(DataPage.SEQUENCE.value)Then
	ErrMSg = ErrMSg & "Sequence must be numeric and not null." & VbCrlf
End If

If Not isNumeric(DataPage.RETRY_COUNT.value) Then
	ErrMSg = ErrMSg & "Please fill in the Retry Count." & VbCrlf
else
	if cint(DataPage.RETRY_COUNT.value) <= 0 then
		ErrMSg = ErrMSg & "Retry Count must be greater than zero." & VbCrlf
	end if
End If

If DataPage.RETRY_WAIT_TIME.value <> "" AND Not IsNumeric(DataPage.RETRY_WAIT_TIME.value)  Then
	ErrMSg = ErrMSg & "Retry Wait Time must be numeric." & VbCrlf
End If

If ErrMSg = "" Then
	window.returnvalue = "SAVED"
	Parent.frames("SeqIFrame").document.frames("TOP").document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
	Parent.frames("SeqIFrame").document.frames("TOP").frmSave.Submit
Else
	msgbox ErrMsg, 0, "FNSDesigner"
End If
End Sub

</script>
</head>
<body BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe id="SeqIFrame" FRAMEBORDER="0" src="TransmissionSeq-f.asp?STATUS=<%= Request.QueryString("STATUS") %>&ROUTING_PLAN_ID=<%= Request.QueryString("ROUTING_PLAN_ID") %>&TRANSMISSION_SEQ_STEP_ID=<%= Request.QueryString("TRANSMISSION_SEQ_STEP_ID") %>" WIDTH="100%" HEIGHT="80%">
</iframe>
<br><br>
<table align="LEFT">
<tr>
<td CLASS="LABEL"><button CLASS="StdButton" <% If MODE="RO" Then Response.Write(" DISABLED ") %> NAME="BtnSave" ACCESSKEY="S"><u>S</u>ave</button></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCancel" ACCESSKEY="C"><U>C</U>lose</button></td>
</tr>
</table>
</body>
</html>