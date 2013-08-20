<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	
	if CStr(Request.QueryString("MODE")) <> "" then	MODE = CStr(Request.QueryString("MODE"))
	
	dim bShowSave, bShowClose
	
	bShowSave = false 
	bShowClose = true
	
	Select Case MODE
		Case "RO"
			bShowSave = false
		Case "RW"
	End Select		

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Attribute Override</title>
<STYLE TYPE="text/css">
HTML {width: 230pt; height: 140pt}
</STYLE>
<SCRIPT LANGUAGE="JScript">
function CAccessPermissionsSearchObj()
{
	this.Selected = false;
}
var AccessPermissionsSearchObj = new CAccessPermissionsSearchObj();

</SCRIPT>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("AccessPermissionsFrame").document.frames.ExeSave();
	AccessPermissionsSearchObj.Selected = true;
</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	AccessPermissionsSearchObj = window.dialogArguments;
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	//if (document.frames("AccessPermissionsFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="OverrideFrame" SRC="Attribute_Override-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
</iframe>
</body>
<TABLE WIDTH="100%">
<TR>
<% if bShowClose = true then %>
<TD CLASS=LABEL ALIGN=RIGHT><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
<%	end if %>
</TR>
</TABLE>
</html>
