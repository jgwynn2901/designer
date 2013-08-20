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
<title>Email Subject and Body Template Files</title>
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
	If DataPage.SUBJECT_FILE.Value = "" and DataPage.BODY_FILE.Value = "" Then
		ErrMSg = "Please enter at least one of email subject or body template file name" & VbCrlf
	End If
	If ErrMsg = "" Then
		Parent.frames("OutputIFrame").document.frames("TOP").frmSave.Submit
	Else
		MsgBox ErrMsg,48,"FNSNetDesigner"
	End If
End Sub

</script>
</head>
<body BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe id="OutputIFrame" FRAMEBORDER="0" src="SubjBody-f.asp?<%= Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
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