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
<title>Account Vendors</title>
<style TYPE="text/css">
HTML {width: 270pt; height: 150pt}
</style>
<script LANGUAGE="JScript">
function CBranchAssignRuleSearchObj()
{
	this.Selected = false;
}
var BranchAssignRuleSearchObj = new CBranchAssignRuleSearchObj();

</script>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("AccVendorAddFrame").document.frames.ExeSave();
	BranchAssignRuleSearchObj.Selected = true;
</script>

<script LANGUAGE="JavaScript" FOR="BtnCopy" EVENT="onclick">
	document.frames("AccVendorAddFrame").document.frames.ExeCopy();
	BranchAssignRuleSearchObj.Selected = true;
</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	BranchAssignRuleSearchObj = window.dialogArguments;
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("AccVendorAddFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="AccVendorAddFrame" SRC="AccVendorAddDetails-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
</iframe>
</body>
<table>
<tr>
<%
	if bShowSave = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSave" ACCESSKEY="S"><u>S</u>ave</button></td>
<%	end if 
	if bShowClose = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnClose">Close</button></td>
<%	end if %>
</tr>
</table>
</html>
