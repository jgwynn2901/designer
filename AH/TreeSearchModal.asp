<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Search</title>
<style TYPE="text/css">
HTML {width: 600pt; height:500pt}
</style>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function BtnSearch_onclick() {
TabFrame.document.frames("TOP").ExeSearch()
}

function BtnGoto_onclick() {
lret = TabFrame.document.frames("WORKAREA").SelectRow()
if (lret != "-1")
{
window.dialogArguments.goto = true;
window.dialogArguments.ahsid = lret;
window.close();
}
else
{
alert ("Please select a row")
}
}

function BtnClear_onclick() {
TabFrame.document.frames("TOP").ClearSearch()
}

function BtnCopy_onclick() {
lret = TabFrame.document.frames("WORKAREA").CopyItem()
	if (lret == "-1")
	{
		alert ("Please Select a row")
	}
}

function BtnClose_onclick() {
window.close()
}

//-->
</SCRIPT>
<script LANGUAGE="JavaScript" FOR="TabsControl" EVENT="onscriptletevent(theEvent,theData)">
		DoOnscriptletevent(theEvent,theData)
</script>


<SCRIPT LANGUAGE="JavaScript">
function OnTabFramesReady()
{

	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);
	AddTab("Business Entity",120,  "AHNodeSearch-f.asp?AHSID=<%= Request.QueryString("AHSID") %>");
	AddTab("Call Flow", 120, "AHCallFlowSearch-f.asp?CONTAINERTYPE=MODAL&AHSID=<%= Request.QueryString("AHSID") %>");
	AddTab("Routing Plan",120,  "AHRoutingPlanSearch-f.asp?AHSID=<%= Request.QueryString("AHSID") %>");
	AddTab("Policy", 120, "AHPolicySearch-f.asp?AHSID=<%= Request.QueryString("AHSID") %>");
	SetActiveTabViaGet("Business Entity");
}
</script>
<SCRIPT SRC='..\LIB\TabFrames.js'></SCRIPT>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" bgcolor='<%= BODYBGCOLOR %>' bottommargin=0 rightmargin=0>
<OBJECT data=..\Scriptlets\TabScriptlet.htm id=TabsControl style="LEFT: 0px; TOP: 0px" 
	type=text/x-scriptlet></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1" HEIGHT="1"></iframe>
<BR>
<TABLE>
<TR>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnGoto" ACCESSKEY="S"  LANGUAGE=javascript onclick="return BtnGoto_onclick()"><u>G</u>oto Node</button></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCopy" ACCESSKEY="P" LANGUAGE=javascript onclick="return BtnCopy_onclick()">Co<U>p</U>y</button></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnClose" ACCESSKEY="O"  LANGUAGE=javascript onclick="return BtnClose_onclick()">Cl<U>o</U>se</button></td>
</TR>
</table>
</body>
</html>