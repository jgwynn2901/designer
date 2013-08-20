
var nCurrTabIndex = 0;
var aTabLabels = new Array();
var aTabURLs = new Array();
var aTabWidths = new Array();
var aTabUsePost = new Array();
var aTabEnabled = new Array();

aTabUsePost[0] = 0;
aTabUsePost[1] = 0;
aTabUsePost[2] = 0;
aTabUsePost[3] = 0;
aTabUsePost[4] = 0;
aTabUsePost[5] = 0;
aTabUsePost[6] = 0;

aTabLabels[0] = "";
aTabLabels[1] = "";
aTabLabels[2] = "";
aTabLabels[3] = "";
aTabLabels[4] = "";
aTabLabels[5] = "";
aTabLabels[6] = "";

aTabWidths[0] = 1;
aTabWidths[1] = 1;
aTabWidths[2] = 1;
aTabWidths[3] = 1;
aTabWidths[4] = 1;
aTabWidths[5] = 1;
aTabWidths[6] = 1;

aTabURLs[0] = "";
aTabURLs[1] = "";
aTabURLs[2] = "";
aTabURLs[3] = "";
aTabURLs[4] = "";
aTabURLs[5] = "";
aTabURLs[6] = "";

aTabEnabled[0] = true;
aTabEnabled[1] = true;
aTabEnabled[2] = true;
aTabEnabled[3] = true;
aTabEnabled[4] = true;
aTabEnabled[5] = true;
aTabEnabled[6] = true;

function AddTab(cTabName, nTabWidth, cTabURL, tabPost, lEnabled)
{
	if (nCurrTabIndex <= 6)
	{
		aTabLabels[nCurrTabIndex] = cTabName;
		aTabURLs[nCurrTabIndex] = cTabURL;
		aTabWidths[nCurrTabIndex] = nTabWidth;
		aTabUsePost[nCurrTabIndex] = tabPost;
		if (lEnabled != null)
			aTabEnabled[nCurrTabIndex] = lEnabled;

		TabsControl.labels(aTabLabels[0], aTabWidths[0], aTabEnabled[0], aTabLabels[1], aTabWidths[1], aTabEnabled[1], aTabLabels[2], aTabWidths[2], aTabEnabled[2], aTabLabels[3],  aTabWidths[3], aTabEnabled[3], aTabLabels[4], aTabWidths[4], aTabEnabled[4], aTabLabels[5], aTabWidths[5], aTabEnabled[5], aTabLabels[6], aTabWidths[6], aTabEnabled[6]);
		nCurrTabIndex++;
	}
}

function enableTab(cTabName)
{
	var nIndex = findTab(cTabName);
	if (nIndex != -1)
		TabsControl.enableTab(nIndex+1);
}

function disableTab(cTabName)
{
	var nIndex = findTab(cTabName);
	if (nIndex != -1)
		TabsControl.disableTab(nIndex+1);
}

function SetActiveTabViaGet(cTabName)
{
	var nIndex = findTab(cTabName);
	if (nIndex != -1)
		{
		TabsControl.setActiveTab(nIndex + 1);
		TabFrame.window.location = aTabURLs[nIndex];
		}
}


function window.onload()
{
	TabsControl.style.posTop = 0;
	TabsControl.style.posLeft = 0;
	TabsControl.style.pixelWidth = document.body.clientWidth;
	TabsControl.style.pixelHeight = 30;
	
	document.all.TabFrame.style.top = TabsControl.style.pixelHeight;
	document.all.TabFrame.style.left = 0;
	document.all.TabFrame.style.width = document.body.clientWidth;
	document.all.TabFrame.style.height = document.body.clientHeight - TabsControl.style.pixelHeight;
			
	OnTabFramesReady();
}

function OnCloseActiveFrame()
{
	window.close();
}

function CanActiveFrameUnloadNow()
{
	if (TabFrame.document.body.getAttribute("CanDocUnloadNowInf") == "YES")
		return TabFrame.CanDocUnloadNow();
	else
		return true;
	
}

function UpdateFrame(nIdx)
{
	if (aTabUsePost[nIdx] == 1)
		TabFrame.PostTo(aTabURLs[nIdx]);
	else
		TabFrame.window.location = aTabURLs[nIdx];
}


function DoOnscriptletevent(theEvent, cData)
{
	var bRet = false;
	
	if ("OnTabClick" != theEvent)
		return false;
	
	switch (cData)
	{
		case "1":
			if (CanActiveFrameUnloadNow() && aTabURLs[0] != "" )
			{
				UpdateFrame(0);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());
			break;

		case "2":
			if (CanActiveFrameUnloadNow() && aTabURLs[1] != "" )
			{
				UpdateFrame(1);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());
			break;

		case "3":
			if (CanActiveFrameUnloadNow() && aTabURLs[2] != "" )
			{
				UpdateFrame(2);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());
			break;

		case "4":
			if (CanActiveFrameUnloadNow() && aTabURLs[3] != "" )
			{
				UpdateFrame(3);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());
		
			break;

		case "5":
			if (CanActiveFrameUnloadNow() && aTabURLs[4] != "" )
			{
				UpdateFrame(4);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());
			break;

		case "6":
			if (CanActiveFrameUnloadNow() && aTabURLs[5] != "" )
			{
				UpdateFrame(5);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());

			break;

		case "7":
			if (CanActiveFrameUnloadNow() && aTabURLs[6] != "" )
			{
				UpdateFrame(6);
				bRet = true;
			}
			else
				TabsControl.setActiveTab(TabsControl.getPreviousTabNum());

			break;

		default:
			break;
	}
	
	return bRet;
}

function findTab(cTabName)
{
var nIndex = 0;
while (nIndex <= 6)
	{
	if (aTabLabels[nIndex] == cTabName)
		break;
	nIndex++;
	}
return ((nIndex <= 6) ? nIndex : -1);
}

function ChangeTabURL(cTabName, cNewURL)
{
	var nIndex = findTab(cTabName);
	if (nIndex != -1)
		aTabURLs[nIndex] = cNewURL;
}
