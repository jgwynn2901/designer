<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<title>Search</title>
</head>
<style>
BODY { background:beige }
</style>
<body LEFTMARGIN="0" TOPMARGIN="0" OnCloseActiveFrameInf=YES>
<OBJECT data="..\Scriptlets\TabScriptlet.htm" id=TabsControl style="LEFT: 0px; TOP: 0px" type=text/x-scriptlet></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1" HEIGHT="1">
</iframe>
</body>
</html>
<SCRIPT SRC='..\LIB\TabFrames.js'></SCRIPT>

<script LANGUAGE="JavaScript" FOR="TabsControl" EVENT="onscriptletevent(theEvent,theData)">
	DoOnscriptletevent(theEvent,theData);
</script>

<SCRIPT LANGUAGE="JavaScript">

function OnTabFramesReady()
{
	AddTab("Node",120,  "NodeSearch-f.asp");
	AddTab("Call Flow", 120, "CallFlowSearch-f.asp");
	AddTab("Routing Plan",120,  "RoutingSearch-f.asp");
	AddTab("Policy", 120, "PolicySearch-f.asp");
	SetActiveTabViaGet("Node");
}

</script>
