<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<% 	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	
	
	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	
	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	If HasModifyPrivilege("FNSD_TIP",SECURITYPRIV) <> True Then MODE = "RO"
	
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	
	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack
	dim bShowSearchTab, bShowDetailsTab
	
	'Response.Write "CONTAINERTYPE: " & CONTAINERTYPE & " </br>"
	'Response.Write "DETAILONLY: " & DETAILONLY & " </br>"
	'Response.Write "AHSID: " & Request.QueryString("AHSID") & " </br>"
	
	bShowNew = true
	bShowCopy = true
	bShowSave = true 
	bShowSelect = true
	bShowClose = true
	bShowBack = true
	bShowSearchTab = true
	bShowDetailsTab = true

	Select Case CONTAINERTYPE
		Case "FRAMEWORK"
			bShowSelect = false
			bShowClose = false
			If Request.QueryString("AHSID") <> "" Then 	bShowBack = false
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
			bShowBack = false
	
			bShowSearchTab = false
		Case "FALSE"
	End Select		


	'If HasAddPrivilege("FNSD_TIP",SECURITYPRIV) <> True Then
		bShowNew = false
		bShowCopy = false
	'End If
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Search</title>
<%	if CONTAINERTYPE = "MODAL" then %>
<style TYPE="text/css">
HTML {width: 380pt; height: 320pt}
</style>
<%	end if %>
<script SRC="..\LIB\TabFrames.js"></script>

<script language="VBScript">
Sub BtnGrfxBack_OnClick()
<% If Request.QueryString("AHSID") <> "" Then %>
	location.href = "..\AH\NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID")%>"
<% End If %>
End Sub
</script>
<script LANGUAGE="JScript">
function SetDisabledButton(btnObj, bDisable)
{
	if (btnObj != null )
		btnObj.disabled = bDisable;
	
}

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

function CBranchAssignTypeSearchObj()
{
	this.BATID = "";
	this.Selected = "";
}
var BranchAssignTypeSearchObj = new CBranchAssignTypeSearchObj();

function UpdateCurrentObjFromScreen()
{
	BranchAssignTypeSearchObj.BATID = document.frames("TabFrame").document.frames.GetBATID();
}


function IsCurrentSelectionValid()
{
	var ATID = document.frames("TabFrame").document.frames.GetATID();
	
	if (ATID == "")
		return false;
	else
		return true;
}	

function OnTabFramesReady()
{
<% If Request.QueryString("AHSID") <> "" Then %>
	document.all.TabsControl.style.width = document.body.clientWidth - 30;
<% End If%>
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);


<%  if CONTAINERTYPE	= "MODAL" then %>
	BranchAssignTypeSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowDetailsTab = true then %>
	AddTab("Details", 120, "TipsDetails-f.asp?MODE=<%=MODE%>&ATID=<%=Request.QueryString("ATID")%>&AHSID=<%=Request.QueryString("AHSID")%>",1);
<%	end if %>

<%	If (Request.QueryString("ATID") = "NEW" Or Request.QueryString("ATID") = "") And bShowSearchTab <> false Then %>
	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
<%	else %>	
	SetActiveTabViaGet("Details");
<%	end if%>

}
</script>




<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnNew_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		ChangeTabURL("Details", "TipsDetails-f.asp?MODE=<%=MODE%>&ATID=NEW");
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

function BtnSave_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (document.frames("TabFrame").document.frames.ExeSave() == true)
		{
			ConfigureButtonsOnSave();
			//UpdateCurrentObjFromScreen();
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
				BranchAssignTypeSearchObj.Selected = true;
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
			BranchAssignTypeSearchObj.Selected = false;
			window.close();
		}
	}
}	


</script>

<script ID="clientEventHandlersVBS" LANGUAGE="VBScript">
Sub Document_OnKeyDown()
	if document.frames("TabFrame").document.readyState = "complete" Then 
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
<% If Request.QueryString("AHSID") <> "" Then %>
<!--#include file="..\lib\NavBack.inc"-->
<% End If %>  
<object VIEWASTEXT data="..\Scriptlets\Tabscriptlet.htm" id="TabsControl" type="text/x-scriptlet"></object>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="100">
</iframe>

<table>
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
</body>
</html>

