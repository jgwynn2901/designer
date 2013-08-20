<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Output Definition Search</title>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	window.dialogArguments.outputdefid = ""
	'document.all.BtnSelect.innerHTML = window.dialogArguments.SelectButtonLabel
	'document.all.BtnSelect.accessKey = window.dialogArguments.SelectButtonAccessKey
End Sub

Sub BtnCancel_OnClick
	window.dialogArguments.outputdefid = ""
	window.close
End Sub

Sub BtnSelect_OnClick
	OUTPUTDEF_ID =  document.frames("ODSearchIFrame").document.frames("WORKAREA").GetDef
	If OUTPUTDEF_ID = -1 OR OUTPUTDEF_ID="X" Then
		MsgBox "Please select a Output Definition.", 0, "FNSNetDesigner"
	Else
		window.dialogArguments.outputdefid = OUTPUTDEF_ID
		window.close	
	End If
End Sub

Sub Document_OnKeyDown()
	If window.event.altKey Then
		KeyPress = Chr(window.event.keyCode)
		Select Case KeyPress
			case "C":
				document.frames("ODSearchIFrame").document.frames("TOP").BtnSearch_onclick
			case "L":
				document.frames("ODSearchIFrame").document.frames("TOP").BtnClear_onclick
		End Select
	End If
End Sub

</script>
</head>

<body BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe id="ODSearchIFrame" FRAMEBORDER="0" src="ODSearch-f.asp" WIDTH="100%" HEIGHT="80%">
</iframe>
<br><br>
<table align="LEFT">
<tr>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSelect" ACCESSKEY="S"><u>S</u>elect</button></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCancel">Cancel</button></td>
</tr>
</table>
</body>
</html>
