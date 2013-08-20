<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	SELECTONLY = "FALSE"
	
	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	if CStr(Request.QueryString("SEARCHONLY")) <> "" then SEARCHONLY = CStr(Request.QueryString("SEARCHONLY"))
	if CStr(Request.QueryString("SELECTONLY")) <> "" then SELECTONLY = CStr(Request.QueryString("SELECTONLY"))

	
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
<title>Vehicle Maintenance</title>
<!--
<STYLE TYPE="text/css">
HTML {HTML {width:700px; height: 750px}
</STYLE>
-->
<SCRIPT LANGUAGE="JScript">
function CVehicleSearchObj()
{
	this.Selected = false;
}
var VehicleSearchObj = new CVehicleSearchObj();

</SCRIPT>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("VehicleFrame").document.frames.ExeSave();
	VehicleSearchObj.Selected = true;
</script>


<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">

	VehicleSearchObj = window.dialogArguments;
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("VehicleFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="VehicleFrame" SRC="VehicleDetails-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="95%" SCROLLING=NO>
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
