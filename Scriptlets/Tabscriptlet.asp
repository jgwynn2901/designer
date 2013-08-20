<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Developer Studio">
<%strConnString=mid(Session("ConnectionString"),1,instr(session("ConnectionString"),";")-1)%>
<title> FNSNet Designer (<%=mid(strConnString,8,len(strConnString)-7)%>)</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script LANGUAGE="JavaScript">

////Globals:

var currentTabNum = -1;
var previousTabNum = -1;
var currentTab;
var tabBase;
var firstFlag = true;


//a public function that the container uses to pass in values for the labels
function public_Labels(label1, width1, label2, width2, label3, width3, label4, width4, label5, width5, label6, width6, label7, width7)
{
	if (label1 != "")
	{
		t1.innerText = label1;
		t1.width = width1;
	}

	if (label2 != "")
	{
		t2.innerText = label2;
		t2.width = width2;
	}

	if (label3 != "")
	{
		t3.innerText = label3;
		t3.width = width3;
	}

	if (label4 != "")
	{
		t4.innerText = label4;
		t4.width = width4;
	}

	if (label5 != "")
	{
		t5.innerText = label5;
		t5.width = width5;
	}

	if (label6 != "")
	{
		t6.innerText = label6;
		t6.width = width6;
	}
	if (label7 != "")
	{
		t7.innerText = label7;
		t7.width = width7;
	}
}


function public_GetPreviousTabNum()
{
	return previousTabNum;
}	

function public_GetCurrentTabNum()
{
	return currentTabNum;
}	


function changeTabs(){

	if(firstFlag == true){
		currentTab = t1;
		tabBase = t1base;
		firstFlag = false;
	}

	if (window.event.srcElement.className == "tab")
	{

		currentTab.className = "tab";
		tabBase.style.backgroundColor = "white";

		currentTab = window.event.srcElement;
		tabBaseID = currentTab.id + "base";
		tabContentID = currentTab.id + "Contents";
		tabBase = document.all(tabBaseID);

		currentTab.className = "selTab";		
		tabBase.style.backgroundColor = "";

		var str = new String(currentTab.id);

		previousTabNum = currentTabNum;
		currentTabNum = str.charAt(1);

		if (!window.external.frozen)
			window.external.raiseEvent("OnTabClick", currentTabNum);

		delete str;
	}
}


function public_SetActiveTab(index)
{

	if(firstFlag == true){
		currentTab = t1;
		tabBase = t1base;
		firstFlag = false;
	}

	var curTabID = "t" + index
	var newTab = document.all(curTabID)

	if ((null != newTab) && (null != currentTab))
	{
		currentTab.className = "tab";

		tabBase.style.backgroundColor = "white";
		currentTab = newTab;
		tabBaseID = currentTab.id + "base";
		tabContentID = currentTab.id + "Contents";
		tabBase = document.all(tabBaseID);
		currentTab.className = "selTab";		
		tabBase.style.backgroundColor = "";

		var str = new String(currentTab.id);

		if (-1 == currentTabNum)
			previousTabNum = str.charAt(1);
		else
			previousTabNum = currentTabNum;

		currentTabNum = str.charAt(1);

		delete str;
	}
	

}
</script>







</head>
<body onclick="changeTabs()" BGCOLOR="#d6cfbd">
<table height="100%" width="100%" CELLPADDING="0" CELLSPACING="0" STYLE="position:absolute; top:4; left:0">
	<tr>
		<td width=2>&nbsp;</td>
		<td ID="t1" CLASS="selTab" HEIGHT="20"></td>
		<td ID="t2" CLASS="tab"></td>
		<td ID="t3" CLASS="tab"></td>
		<td ID="t4" CLASS="tab"></td>
		<td ID="t5" CLASS="tab"></td>
		<td ID="t6" CLASS="tab"></td>
		<td ID="t7" CLASS="tab"></td>
	</tr>
	<tr>
		<td STYLE="height:1; background-color:white"></td>
		<td ID="t1base" STYLE="height:1; border-left:solid thin white"></td>
		<td ID="t2base" STYLE="height:1; background-color:white"></td>
		<td ID="t3base" STYLE="height:1; background-color:white"></td>
		<td ID="t4base" STYLE="height:1; background-color:white"></td>
		<td ID="t5base" STYLE="height:1; background-color:white"></td>
		<td ID="t6base" STYLE="height:1; background-color:white"></td>
		<td ID="t7base" STYLE="height:1; background-color:white"></td>	
	</tr>
	<td span=8>&nbsp;</td>
	<tr>
	</tr>
</table>
</body>
</html>
