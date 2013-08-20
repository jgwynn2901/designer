<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<title>Policy Jurisdiction State Maintenance</title>
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
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<STYLE TYPE="text/css"> HTML {width: 290pt; height: 330pt}</STYLE>	
<SCRIPT LANGUAGE="JScript">
function CJurisdictionObj()
{
	this.Selected = false;
}
var JurisdictionObj = new CJurisdictionObj();
</SCRIPT>
<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("StateFrame").document.frames.ExeSave();
	JurisdictionObj.Selected = true;
</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	JurisdictionObj = window.dialogArguments;
</script>
<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("StateFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</HEAD>
<body LEFTMARGIN="0" TOPMARGIN="0" RIGHTMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="StateFrame" SRC="PolicyJurisStateDetails-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="90%" scrolling="no"></iframe>
<TABLE>
<TR>
<%	if bShowSave = true then %>
	<TD CLASS="LABEL"><BUTTON CLASS="StdButton" NAME="BtnSave" ACCESSKEY="S" id="BUTTON1" type="button"><U>S</U>ave</BUTTON></TD>
<%	end if 
if bShowClose = true then %>
	<TD CLASS="LABEL"><BUTTON CLASS="StdButton" NAME="BtnClose" type="button">Close</BUTTON></TD>
<%	end if %>
</TR>
</TABLE>
</body>
</HTML>
