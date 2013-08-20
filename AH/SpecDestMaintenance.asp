<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<%
	dim cMode
	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack
	dim bShowSearchTab, bShowDetailsTab
	dim cContainerType, cDetailOnly
	
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	cMode = "RW"
	cDetailOnly = "FALSE"
	cContainerType = "MODAL"

	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	If HasModifyPrivilege("FNSD_SPECIFIC_DESTINATION",SECURITYPRIV) <> True Then 
		cMode = "RO"
	end if
	if CStr(Request.QueryString("DETAILONLY")) <> "" then 
		cDetailOnly = CStr(Request.QueryString("DETAILONLY"))
	end if
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then 
		cContainerType = CStr(Request.QueryString("CONTAINERTYPE"))
	end if
		
	bShowNew = true
	bShowCopy = true
	bShowSave = true 
	bShowSelect = true
	bShowClose = true
	bShowBack = true
	bShowSearchTab = true
	bShowDetailsTab = true

	Select Case cContainerType
		Case "FRAMEWORK"
			bShowSelect = false
			bShowNew = false
			bShowClose = false
		Case "MODAL"
			bShowBack = false
		Case "DIALOG"
	End Select		

	Select Case cMode
		Case "RO"
			bShowNew = false
			bShowCopy = false
			bShowSave = false
		Case "RW"
	End Select		

	Select Case cDetailOnly
		Case "TRUE"
			bShowNew = false
			bShowCopy = false
			bShowSelect = false
			
			bShowSearchTab = false
		Case "FALSE"
	End Select		
	
	If HasAddPrivilege("FNSD_SPECIFIC_DESTINATION",SECURITYPRIV) <> True Then
		bShowNew = false
		bShowCopy = false
	End If

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Search</title>
<SCRIPT SRC='..\LIB\TabFrames.js'></SCRIPT>

<SCRIPT LANGUAGE="JScript">
function SetHiddenButton(btnObj, bHide)
{
	if (btnObj != null )
	{
		if (bHide == true)
		{
			btnObj.className = "StdButtonHidden";
			btnObj.tabIndex = -1;
		}
		else
		{
			btnObj.className = "StdButton";
			btnObj.tabIndex = 0;
		}
	}
}
function SetDisabledButton(btnObj, bDisable)
{
	if (btnObj != null )
		btnObj.disabled = bDisable;
	
}

function CCarrierSearchObj()
{
	this.COID = "";
	this.COIDName = "";
	this.Selected = false;	
}
var CarrierSearchObj = new CCarrierSearchObj();

function ConfigureButtonsOnSave()
{
	SetDisabledButton(document.all.BtnCopy,false);
	SetDisabledButton(document.all.BtnNew,false);
	SetDisabledButton(document.all.BtnSelect,false);
}

function ConfigureButtonsOnNew()
{
	SetDisabledButton(document.all.BtnSave,false);
	SetDisabledButton(document.all.BtnCopy, true);
	SetDisabledButton(document.all.BtnNew, true);
	SetDisabledButton(document.all.BtnSelect, true);
}

function ConfigureButtonsOnActivateTab(inTab)
{
	if (inTab == "Search")
	{
		SetDisabledButton(document.all.BtnSave,true);
		SetDisabledButton(document.all.BtnCopy,true);
		SetDisabledButton(document.all.BtnBack,true);
		SetDisabledButton(document.all.BtnNew,false);
		SetDisabledButton(document.all.BtnSelect,false);

		SetHiddenButton(document.all.BtnSave,true);
		SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnBack,true);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);

	}
	else if (inTab == "Details")
	{
		if (IsCurrentSelectionValid()== false)
		{
			SetDisabledButton(document.all.BtnSave,true);
			SetDisabledButton(document.all.BtnCopy,true);
		}
		else
		{
			SetDisabledButton(document.all.BtnSave,false);
			SetDisabledButton(document.all.BtnCopy,false);
		}
		SetDisabledButton(document.all.BtnSelect,false);
		SetDisabledButton(document.all.BtnBack,false);
		SetDisabledButton(document.all.BtnNew,false);
		
		SetHiddenButton(document.all.BtnSave,false);
		SetHiddenButton(document.all.BtnCopy,false);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnBack,false);

	}
	
}


function UpdateCurrentObjFromScreen()
{
	CarrierSearchObj.COID = document.frames("TabFrame").document.frames.GetCOID();
	CarrierSearchObj.COIDName = document.frames("TabFrame").document.frames.GetCOIDName();
}

function IsCurrentSelectionValid()
{
	var SDID = document.frames("TabFrame").document.frames.GetSDID();
	if (SDID == "")
		return false;
	else
		return true;
}	

function OnTabFramesReady()
{
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight);
<%  if cContainerType	= "MODAL" then %>
	CarrierSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowSearchTab = true then %>
	AddTab("Search",120, "AHSpecDestSearch-f.asp?MODE=<%=cMode%>&FULLSEARCH", 1);
<%	end if 
	if bShowDetailsTab = true then %>
	AddTab("Details", 120, "SpecDestDetails-f.asp?MODE=<%=cMode%>&FULLSEARCH",1);
<%	end if %>

	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
}
</SCRIPT>

<script LANGUAGE="JavaScript" FOR="TabsControl" EVENT="onscriptletevent(theEvent,theData)">
	if (theData == "1") othData = "2";
	else othData = "1";	

	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (DoOnscriptletevent(theEvent,theData) == true)
		{
			if (theData == "1")//search
				ConfigureButtonsOnActivateTab("Search");
			else
				ConfigureButtonsOnActivateTab("Details");
		}
	}
	else
		document.all.TabsControl.SetActiveTab(othData);		
</script>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnNew_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		ChangeTabURL("Details", "ContactDetails-f.asp?MODE=<%=MODE%>&COID=NEW");
		SetActiveTabViaGet("Details");
		ConfigureButtonsOnActivateTab("Details");
		ConfigureButtonsOnNew();
	}		
}
function BtnBack_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		document.all.TabsControl.SetActiveTab("1");
		if (DoOnscriptletevent("OnTabClick", "1") == true)
			ConfigureButtonsOnActivateTab("Search");
	}			
}

function BtnCopy_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (document.frames("TabFrame").document.frames.IsDirty() == true)
			alert("Data has changed. You must save your changes or reselect the item from the search screen before choosing Copy.");
		else
			document.frames("TabFrame").document.frames.ExeCopy();
	}			
}
function BtnSave_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (document.frames("TabFrame").document.frames.ExeSave() == true)
		{
			ConfigureButtonsOnSave();
			UpdateCurrentObjFromScreen();
		}
	}
}

function BtnSelect_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (IsCurrentSelectionValid() == false)
			alert("Nothing to select.");
		else	
		{
			if (document.frames("TabFrame").document.frames.IsDirty() == true)
				alert("Data has changed. You must save your changes or reselect the item from the search screen before choosing Select.");
			else
			{
				UpdateCurrentObjFromScreen();
				CarrierSearchObj.Selected = true;
				window.close();
			}
		}
	}
}

function BtnClose_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (CanActiveFrameUnloadNow() == true)
		{	
			CarrierSearchObj.Selected = false;
			window.close();
		}
	}		
}
</SCRIPT>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE="VBScript">

Sub Document_OnKeyDown()
	if document.frames("TabFrame").document.readyState = "complete" then
		if document.frames("TabFrame").document.frames(0).name = "TOP" then
			If window.event.altKey Then
				KeyPress = Chr(window.event.keyCode)
				Select Case KeyPress
					case "H":
						document.frames("TabFrame").document.frames("TOP").BtnSearch_onclick
					case "L":
						document.frames("TabFrame").document.frames("TOP").BtnClear_onclick
				End Select
			End If
		End if
	End If		
End Sub

</SCRIPT>

</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<OBJECT data=..\Scriptlets\TabScriptlet.htm id=TabsControl style="LEFT: 0px; TOP: 0px" 
	type=text/x-scriptlet VIEWASTEXT></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="10">
</iframe>
</body>
</html>
