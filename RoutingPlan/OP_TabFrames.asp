<!--#include file="..\lib\common.inc"-->
<%

Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<title>Tab Frames</title>
</head>
<style>
BODY { background:#d6cfbd }
</style>
<BODY topmargin="0" rightmargin="0" leftmargin="0" bottommargin="0" OnCloseActiveFrameInf = YES>
<OBJECT data="../Scriptlets/TabScriptlet.asp" id=TabsControl style="LEFT: 0px; TOP: 0px" type=text/x-scriptlet VIEWASTEXT></OBJECT>
<DIV>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="100" HEIGHT="100"></iframe>
</DIV>
</body>
</html>
<script LANGUAGE="JavaScript">

var curTabIndex = 0;
var tabLabels = new Array();
var tabURLs = new Array();
var tabWidths = new Array();

tabLabels[0] = "";
tabLabels[1] = "";
tabLabels[2] = "";
tabLabels[3] = "";
tabLabels[4] = "";
tabLabels[5] = "";
tabLabels[6] = "";

tabWidths[0] = 1;
tabWidths[1] = 1;
tabWidths[2] = 1;
tabWidths[3] = 1;
tabWidths[4] = 1;
tabWidths[5] = 1;
tabWidths[6] = 1;

tabURLs[0] = "";
tabURLs[1] = "";
tabURLs[2] = "";
tabURLs[3] = "";
tabURLs[4] = "";
tabURLs[5] = "";
tabURLs[6] = "";


function AddTab(tabName, tabWidth, tabURL)
{
	if (curTabIndex <= 6)
	{
		tabLabels[curTabIndex] = tabName;
		tabURLs[curTabIndex] = tabURL;
		tabWidths[curTabIndex] = tabWidth;

		TabsControl.Labels(tabLabels[0], tabWidths[0], tabLabels[1], tabWidths[1], tabLabels[2], tabWidths[2], tabLabels[3],  tabWidths[3], tabLabels[4], tabWidths[4], tabLabels[5], tabWidths[5], tabLabels[6], tabWidths[6]);
		curTabIndex = curTabIndex + 1;
	}
}

function SetActiveTab(tabName)
{
	var bNotMatch = true;
	var index = 0;
	while ((bNotMatch) && (index <= 6))
	{
		if (tabLabels[index] == tabName)
		{
			bNotMatch = false;
			TabsControl.SetActiveTab(index + 1);
			
			TabFrame.document.location = tabURLs[index];
		}
		index = index + 1;
	}
}


function OnCloseActiveFrame()
{
	parent.window.close();
}

function CanActiveFrameUnloadNow()
{
	if (TabFrame.document.readyState == "complete")
	{
		if (TabFrame.document.body.getAttribute("CanDocUnloadNowInf")=="YES")
			return TabFrame.CanDocUnloadNow();
		else
			return true;
	}
}
function ReCalcLayout()
{

	TabsControl.style.pixelTop = 0;
	TabsControl.style.pixelLeft = 0;
	TabsControl.style.pixelWidth = document.body.clientWidth;
	TabsControl.style.pixelHeight = 30;
	
	document.all.TabFrame.top = TabsControl.style.pixelHeight + 1;
	document.all.TabFrame.style.pixelLeft = 0;
	document.all.TabFrame.style.pixelWidth = document.body.clientWidth;
	document.all.TabFrame.style.pixelHeight = document.body.clientHeight - TabsControl.style.pixelHeight;

}
function window.onload()
{
ReCalcLayout();
parent.OnTabFramesReady();
}


function window.onresize()
{
ReCalcLayout();
}
</script>


<script LANGUAGE="JavaScript" FOR="TabsControl" EVENT="onscriptletevent(theEvent,theData)">

	if ("OnTabClick" != theEvent)
		return;

	switch (theData)
	{
		case "1":
			if (CanActiveFrameUnloadNow() && tabURLs[0] != "")
				TabFrame.window.location = tabURLs[0];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());
			break;

		case "2":
			if (CanActiveFrameUnloadNow() && tabURLs[1] != "")
				TabFrame.window.location = tabURLs[1];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());
			break;

		case "3":
			if (CanActiveFrameUnloadNow() && tabURLs[2] != "")
				TabFrame.window.location = tabURLs[2];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());
			break;

		case "4":
			if (CanActiveFrameUnloadNow() && tabURLs[3] != "")
				TabFrame.window.location = tabURLs[3];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());
		
			break;

		case "5":
			if (CanActiveFrameUnloadNow() && tabURLs[4] != "")
				TabFrame.window.location = tabURLs[4];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());
			break;

		case "6":
			if (CanActiveFrameUnloadNow() && tabURLs[5] != "")
				TabFrame.window.location = tabURLs[5];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());

			break;

		case "7":
			if (CanActiveFrameUnloadNow() && tabURLs[6] != "")
				TabFrame.window.location = tabURLs[6];
			else
				TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());

			break;

		default:
			break;
	}
 

</script>