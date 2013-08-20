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
<title>File Output</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub BtnCancel_OnClick
MsgRet = "1"
	window.close
End Sub

Sub BtnSave_onclick
	Set DataPage = Parent.frames("OutputIFrame").document.frames("TOP").document.all
	If DataPage.CR_FILE.Value = "" Then
		ErrMSg = "Please enter the Crystal Reports Filename." & VbCrlf
	End If
	If DataPage.OUTPUT_FILE.Value = "" Then
		ErrMSg = ErrMSg & "Please enter the output Filename." & VbCrlf
	End If
	If DataPage.RULE_ID.innerText = "" then
		ErrMSg = ErrMSg & "The Enabling Rule is required."
	end if
	If ErrMsg = "" Then
		Parent.frames("OutputIFrame").document.frames("TOP").frmSave.Submit
	Else
		MsgBox ErrMsg,48,"FNSNetDesigner"
	End If
End Sub

</script>
</head>
<body BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe id="OutputIFrame" FRAMEBORDER="0" src="FileOutPut-f.asp?<%= Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
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