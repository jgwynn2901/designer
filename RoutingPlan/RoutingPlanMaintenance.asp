<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	
	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	SELECTONLY = "FALSE"

	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("SELECTONLY")) <> "" then SELECTONLY = CStr(Request.QueryString("SELECTONLY"))

	If SELECTONLY <> "TRUE" Then
		If HasModifyPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then MODE = "RO"
	Else
		MODE ="RO"
	End If
		

		
	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack
	dim bShowSearchTab
	bShowNew = true
	bShowCopy = true
	bShowSave = true 
	bShowSelect = true
	bShowClose = true
	bShowBack = true
	bShowSearchTab = true

	Select Case CONTAINERTYPE
		Case "FRAMEWORK"
			bShowSelect = false
			bShowClose = false
		Case "MODAL"
			bShowBack = false
		Case "DIALOG"
	End Select		

	Select Case MODE
		Case "RO"
			bShowNew = false
			bShowCopy = false
			bShowSave = false
		Case "RW"
	End Select		

	Select Case DETAILONLY
		Case "TRUE"
			bShowNew = false
			bShowCopy = false
			bShowSelect = false
			
			bShowSearchTab = false
		Case "FALSE"
	End Select		


	If HasAddPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then
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

function CRoutingPlanSearchObj()
{
	this.RPID = "";
	this.RPDesc = "";
	this.Selected = false;
}
var RoutingPlanSearchObj = new CRoutingPlanSearchObj();

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
}

function UpdateCurrentObjFromScreen()
{
	RoutingPlanSearchObj.RPID = document.frames("TabFrame").document.frames.GetRPID();
	RoutingPlanSearchObj.RPDesc = document.frames("TabFrame").document.frames.GetRPDesc();
}

function IsCurrentSelectionValid()
{
	var RPID = document.frames("TabFrame").document.frames.GetRPID();
	
	if (RPID == "")
		return false;
	else
		return true;
}	

function OnTabFramesReady()
{
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);
	
	RoutingPlanSearchObj = window.dialogArguments;
	AddTab("Search",120, "RoutingPlanSearch-f.asp?CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>&ahsid=<%= Request.QueryString("SearchAHSID") %>", 1);

	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
}

</SCRIPT>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnNew_onclick()
{
}

function BtnBack_onclick()
{
}

function BtnCopy_onclick()
{
}

function BtnSave_onclick()
{
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
				RoutingPlanSearchObj.Selected = true;
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
			RoutingPlanSearchObj.Selected = false;
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
<OBJECT data="..\Scriptlets\TabScriptlet.htm" id=TabsControl style="LEFT: 0px; TOP: 0px" type=text/x-scriptlet VIEWASTEXT></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="1000">
</iframe>
</body>
<TABLE>
<%	if bShowNew = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnNew" ACCESSKEY="N" LANGUAGE=javascript onclick="return BtnNew_onclick()"><u>N</u>ew</button></td>
<%	end if
	if bShowCopy = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCopy" ACCESSKEY="C" LANGUAGE=javascript onclick="return BtnCopy_onclick()">Make <u>C</u>opy</button></td>
<%	end if	
	if bShowSave = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSave" ACCESSKEY="S" LANGUAGE=javascript onclick="return BtnSave_onclick()"><u>S</u>ave</button></td>
<%	end if 
	if bShowSelect = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSelect" ACCESSKEY="T" LANGUAGE=javascript onclick="return BtnSelect_onclick()">Selec<u>t</u></button></td>
<%	end if 
	if bShowClose = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnClose" LANGUAGE=javascript onclick="return BtnClose_onclick()">Close</button></td>
<%	end if 
	if bShowBack = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnBack" LANGUAGE=javascript onclick="return BtnBack_onclick()">Back</button></td>
<%	end if %>
</table>
</html>
