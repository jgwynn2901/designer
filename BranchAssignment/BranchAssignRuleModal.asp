<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	
	if CStr(Request.QueryString("MODE")) <> "" then	MODE = CStr(Request.QueryString("MODE"))
	
	dim bShowSave, bShowClose, bShowCopy
	
	bShowSave = true 
	bShowClose = true
	bShowCopy = true
	
	Select Case MODE
		Case "RO"
			bShowSave = false
			bShowCopy = false
		Case "RW"
	End Select		

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Branch Assignment Rule Maintenance</title>
<STYLE TYPE="text/css">
HTML {width: 350pt; height: 175pt}
</STYLE>
<SCRIPT LANGUAGE="JScript">
function CBranchAssignRuleSearchObj()
{
	this.Selected = false;
}
var BranchAssignRuleSearchObj = new CBranchAssignRuleSearchObj();

</SCRIPT>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("BranchAssignRuleFrame").document.frames.ExeSave();
	BranchAssignRuleSearchObj.Selected = true;
</script>

<script LANGUAGE="JavaScript" FOR="BtnCopy" EVENT="onclick">
	document.frames("BranchAssignRuleFrame").document.frames.ExeCopy();
	BranchAssignRuleSearchObj.Selected = true;
</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	BranchAssignRuleSearchObj = window.dialogArguments;
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("BranchAssignRuleFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="BranchAssignRuleFrame" SRC="BranchAssignRuleDetails-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
</iframe>
</body>
<TABLE>
<TR>
<%	if bShowCopy = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCopy" ACCESSKEY="C" LANGUAGE=javascript onclick="return BtnCopy_onclick()">Make <u>C</u>opy</button></td>
<%	end if	
	if bShowSave = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSave ACCESSKEY="S"><U>S</U>ave</BUTTON></TD>
<%	end if 
	if bShowClose = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
<%	end if %>
</TR>
</TABLE>
</html>
