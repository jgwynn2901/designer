<!--#include file="..\lib\common.inc"-->

<html>

<head>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<STYLE TYPE="text/css">
HTML {width: 330pt; height: 200pt}
</STYLE>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

Sub window_onload
	TabFrame.document.location = "ZIPLookup-f.asp?<%=Request.QueryString%>"
End Sub

Sub BtnSelect_onclick
dim oTableObj, cIndex, oRow

set oTableObj = document.frames("TabFrame").document.frames("WORKAREA")
cIndex = oTableObj.getselectedindex( oTableObj.Document.all.tblResult ) 
if cIndex <> "-1" AND cIndex <> "X" AND cIndex <> "null" then
	set oRow = oTableObj.Document.all.tblResult.rows(Cint(cIndex))
	window.dialogArguments.City = oRow.cells(0).innerText
	window.dialogArguments.State = oRow.cells(1).innerText
	window.dialogArguments.Zip = oRow.cells(2).innerText
	window.dialogArguments.County = oRow.cells(3).innerText
	window.dialogArguments.FIPS = oRow.cells(4).innerText
	window.dialogArguments.Country = oRow.cells(5).innerText
	window.close
Else
	msgbox "Nothing selected."
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
<IFRAME FRAMEBORDER="0" ID="TabFrame" WIDTH="100%" HEIGHT="88%" SRC="" >
</IFRAME><br>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelect>Select</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel>Cancel</BUTTON></TD>
</TR>
</TABLE>
</body>
</html>
