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
<title>Fraud Detection Rule Maintenance</title>
<STYLE TYPE="text/css">
HTML {width: 350pt; height: 175pt}
</STYLE>
<SCRIPT LANGUAGE="JScript">
function CFraudDetRuleSearchObj()
{
	this.Selected = false;
}
var FraudDetRuleSearchObj = new CFraudDetRuleSearchObj();

</SCRIPT>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("FraudDetRuleFrame").document.frames.ExeSave();
	FraudDetRuleSearchObj.Selected = true;
</script>


<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	FraudDetRuleSearchObj = window.dialogArguments;
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("FraudDetRuleFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="FraudDetRuleFrame" SRC="FraudDetRuleDetails-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="80%">
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
