<!--#include file="..\Lib\Common.inc"-->
<%
response.expires=0
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT SRC='..\LIB\TabFrames.js'></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function OnTabFramesReady()
{
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);
	AddTab("Contents",120, "HelpTree.asp", 1);
	AddTab("Index", 120, "HelpContents.asp",1);
	SetActiveTabViaGet("Contents");
}
</SCRIPT>
<script LANGUAGE="JavaScript" FOR="TabsControl" EVENT="onscriptletevent(theEvent,theData)">
	if (theData == "2")
	{
	TabFrame.location.href = "HelpContents.asp"
	}
	else
	{
	TabFrame.location.href = "HelpTree.asp"
	}	
</script>
</HEAD>
<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<OBJECT data=..\Scriptlets\TabScriptlet.htm id=TabsControl style="LEFT: 0px; TOP: 0px" 
	type=text/x-scriptlet VIEWASTEXT></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="1000">
</iframe>

</BODY>
</HTML>
