<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	
	if CStr(Request.QueryString("MODE")) <> "" then	MODE = CStr(Request.QueryString("MODE"))
	
	dim bShowSave, bShowClose
	
	bShowSave = true 
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
<title>User Location</title>
<!--<STYLE TYPE="text/css">
HTML {width: 240pt; height: 140pt}
</STYLE>-->
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
	if (document.frames("AccessPermissionsFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="AccessPermissionsFrame" SRC="AccessLocations-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="95%">
</iframe>
</body>
<TABLE>
<TR>
<%	if bShowSave = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSave ACCESSKEY="S"><U>S</U>ave</BUTTON></TD>
<%	end if 
	if bShowClose = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
<%	end if %>
</TR>
</TABLE>
</html>
