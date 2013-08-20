<!--#include file="..\lib\common.inc"-->

<html>

<head>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

Sub window_onload
	TabFrame.document.location = "AHSpecDestSearch-f.asp?CLIENT_NODE=<%=Request.QueryString("CLIENT_NODE")%>"
End Sub

Dim routing_plan_id, ahsid 

Sub BtnSelect_onclick
dim oTableObj, cIndex, oRow

set oTableObj = document.frames("TabFrame").document.frames("WORKAREA")
	cIndex = oTableObj.getselectedindex( oTableObj.Document.all.tblResult ) 
	if cIndex <> "-1" AND cIndex <> "X" AND cIndex <> "null" then
		set oRow = oTableObj.Document.all.tblResult.rows(Cint(cIndex))
		window.dialogArguments.selectedID = oRow.getAttribute("SD")
		window.close
	Else
		msgbox "Please select a Specific Destination."
	End If
End Sub

Sub BtnCancel_onclick
	window.close
End Sub

-->
</script>

<title></title>
</head>

<body leftmargin="0" topmargin="0" bottommargin="0" rightmargin="0" BGCOLOR='<%=BODYBGCOLOR%>'> 
<IFRAME FRAMEBORDER="0" ID="TabFrame" WIDTH="100%" HEIGHT="90%" SRC="" >
</IFRAME><br>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelect>Select</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel>Cancel</BUTTON></TD>
</TR>
</TABLE>
</body>
</html>
