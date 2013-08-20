<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\NavigateBack.inc"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	
	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	If HasAutomaticSecurityPrivilege() = False Then
		If HasModifyPrivilege("FNSD_USERS",SECURITYPRIV) <> True Then MODE = "RO"
	End If

	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	if CStr(Request.QueryString("SEARCHONLY")) <> "" then SEARCHONLY = CStr(Request.QueryString("SEARCHONLY"))
	
	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack
	dim bShowSearchTab, bShowDetailsTab
	
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
	Select Case SEARCHONLY
		Case "TRUE"
			bShowDetailsTab = false

			bShowNew = false
			bShowCopy = false
		Case "FALSE"
	End Select

	If HasAutomaticSecurityPrivilege() = False Then
		If HasAddPrivilege("FNSD_USERS",SECURITYPRIV) <> True Then
			bShowNew = false
			bShowCopy = false
		End If
	End If



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
	location.href = "..\AH\NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID")%>&DROPDOWN=USER"
<% End If %>
End Sub
</script>
<script LANGUAGE="JScript">
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

function CUserSearchObj()
{
	this.UID = "";
	this.UIDName = "";
	this.Selected = "";
}
var UserSearchObj = new CUserSearchObj();


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
			if (IsCurrentSelectionNew()==true)
				SetDisabledButton(document.all.BtnCopy,true);
			else
				SetDisabledButton(document.all.BtnCopy,false);
		}
		
		SetDisabledButton(document.all.BtnSelect,false);
		SetDisabledButton(document.all.BtnBack,false);
		SetDisabledButton(document.all.BtnNew,false);
		
		SetHiddenButton(document.all.BtnSave,false);
		SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnNew,true);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnBack,true);

	}
	else if (inTab == "Accounts")
	{
		SetHiddenButton(document.all.BtnSave,true);
		SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnNew,true);
		SetHiddenButton(document.all.BtnBack,true);
	}
	else if (inTab == "Location")
	{
		SetHiddenButton(document.all.BtnSave,true);
		SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnNew,true);
		SetHiddenButton(document.all.BtnBack,true);
	}
	
}

function UpdateCurrentObjFromScreen()
{
	UserSearchObj.UID = document.frames("TabFrame").document.frames.GetUID();
	UserSearchObj.UIDName = document.frames("TabFrame").document.frames.GetUIDName();
}

function IsCurrentSelectionValid()
{
	var UID = document.frames("TabFrame").document.frames.GetUID();
	
	if (UID == "")
		return false;
	else
		return true;
}	
function IsCurrentSelectionNew()
{
	var UID = document.frames("TabFrame").document.frames.GetUID();
	
	if (UID == "NEW")
		return true;
	else
		return false;
}	

function OnTabFramesReady()
{
<% If Request.QueryString("AHSID") <> "" Then %>
	document.all.TabsControl.style.width = document.body.clientWidth - 30;
<% End If%>
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);


<%  if CONTAINERTYPE	= "MODAL" then %>
	UserSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowSearchTab = true then %>
	AddTab("Search", 60, "UserSearch-f.asp?MODE=<%=MODE%>", 1);
<%	end if 
	if bShowDetailsTab = true then %>
	AddTab("Details", 70, "UserDetails-f.asp?MODE=<%=MODE%>&UID=<%=Request.QueryString("UID")%>&AHSID=<%=Request.QueryString("AHSID")%>",1, false);
	AddTab("Groups", 60, "UserGroups-f.asp?MODE=<%=MODE%>&UID=<%=Request.QueryString("UID")%>&AHSID=<%=Request.QueryString("AHSID")%>",1, false);
	AddTab("Permissions", 90, "UserPermissions-f.asp?MODE=<%=MODE%>&UID=<%=Request.QueryString("UID")%>",1, false);
	AddTab("Accounts", 80, "UserAccounts-f.asp?MODE=<%=MODE%>&UID=<%=Request.QueryString("UID")%>",1, false);
	AddTab("Locations", 80, "UserLocations-f.asp?MODE=<%=MODE%>&UID=<%=Request.QueryString("UID")%>&AHSID=<%=Request.QueryString("AHSID")%>",1, false);
<%	end if %>



<%	If (Request.QueryString("UID") = "NEW" Or Request.QueryString("UID") = "") And bShowSearchTab <> false Then %>
	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
<%	else %>	
	SetActiveTabViaGet("Details");
<%	end if%>

}
</script>

<script LANGUAGE="JavaScript" FOR="TabsControl" EVENT="onscriptletevent(theEvent,theData)">	
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (DoOnscriptletevent(theEvent,theData) == true)
		{
<%	'if DETAILONLY is true there is no Search tab
	if DETAILONLY <> "TRUE" then %>
			if (theData == "1")//search
				ConfigureButtonsOnActivateTab("Search");
			else
<%	end if %>
				ConfigureButtonsOnActivateTab("Details");
		}
	}
	else
		document.all.TabsControl.SetActiveTab(TabsControl.GetPreviousTabNum());		
</SCRIPT>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnNew_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		ChangeTabURL("Details", "UserDetails-f.asp?MODE=<%=MODE%>&UID=NEW");
		enableTab("Details");
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
		else if (IsCurrentSelectionNew() == true)
			alert("You must save the current item before selecting.");
		else	
		{
			if (document.frames("TabFrame").document.frames.IsDirty() == true)
				alert("Data has changed. You must save your changes or reselect the item from the search screen before choosing Select.");
			else
			{
				UpdateCurrentObjFromScreen();
				UserSearchObj.Selected = true;
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
			UserSearchObj.Selected = false;
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

sub VBStop
stop
end sub
</SCRIPT>
</head>


<body LEFTMARGIN="0" RIGHTMARGIN=0  TOPMARGIN=0  BGCOLOR="<%=BODYBGCOLOR%>">
<% If Request.QueryString("AHSID") <> "" Then %>
<!--#include file="..\lib\NavBack.inc"-->
<% End If %>  
<object VIEWASTEXT data="..\Scriptlets\TabScriptlet.htm" id="TabsControl" style="LEFT: 0px; TOP: 0px" type="text/x-scriptlet"></object>
<iframe ID="TabFrame" WIDTH="1000" HEIGHT="1000">
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
