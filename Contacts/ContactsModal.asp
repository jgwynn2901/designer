<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	dim bShowSave, bShowClose
	bShowSave = true 
	bShowClose = true
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Contact Maintenance</title>
<STYLE TYPE="text/css">
HTML {width: 450pt; height: 275pt}
</STYLE>
<SCRIPT LANGUAGE="JScript">
function CContactDataObj()
{
	this.Selected = false;
}
var ContactDataObj = new CContactDataObj();

</SCRIPT>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("ContactFrame").document.frames.ExeSave();
	ContactDataObj.Selected = true;
</script>


<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	ContactDataObj = window.dialogArguments;
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="ContactFrame" SRC="ContactDetails-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
</iframe>
</body>
<TABLE>
<TR>

<%		
	if bShowSave = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSave ACCESSKEY="S"><U>S</U>ave</BUTTON></TD>
<%	end if 
	if bShowClose = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
<%	end if %>
</TR>
</TABLE>
</html>
