<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->


<%

	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	
	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	If HasModifyPrivilege("FNSD_COVERAGE_CODE_XREF",SECURITYPRIV) <> True Then MODE = "RO"

	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	
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


	If HasAddPrivilege("FNSD_COVERAGE_CODE_XREF",SECURITYPRIV) <> True Then
		bShowNew = false
		bShowCopy = false
	End If
	If HasDeletePrivilege("FNSD_COVERAGE_CODE_XREF",SECURITYPRIV) <> True Then
		bShowDelete = false
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
'Sub BtnGrfxBack_OnClick()
'<% 'If Request.QueryString("AHSID") <> "" Then %>
	'location.href = "..\AH\NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID")%>&DROPDOWN=MCBRANCH"
'<% 'End If %>
'End Sub
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

function CoverageSearchObj()
{
	this.XREFID = "";
	this.Selected = "";
}
var CoverageSearchObj = new CoverageSearchObj();


function ConfigureButtonsOnSave()
{
	SetDisabledButton(document.all.BtnCopy,false);
	SetDisabledButton(document.all.BtnNew,false);
	SetDisabledButton(document.all.BtnSelect,false);
}

function ConfigureButtonsOnNew()
{
	SetDisabledButton(document.all.BtnSave,false);
	SetDisabledButton(document.all.BtnCopy,false);
	SetDisabledButton(document.all.BtnNew, true);
	SetDisabledButton(document.all.BtnSelect, true);
}

function ConfigureButtonsOnActivateTab(inTab)
{
	
	if (inTab == "Search")
	{
		SetDisabledButton(document.all.BtnSave,true);
		//SetDisabledButton(document.all.BtnCopy,false);
		SetDisabledButton(document.all.BtnBack,true);
		SetDisabledButton(document.all.BtnNew,false);
		SetDisabledButton(document.all.BtnSelect,false);

		SetHiddenButton(document.all.BtnSave,true);
		//SetHiddenButton(document.all.BtnCopy,false);
		SetHiddenButton(document.all.BtnBack,true);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);


	}
	else if (inTab == "Details")
	{
		if (IsCurrentSelectionValid()== false)
		{
			SetDisabledButton(document.all.BtnSave,false);
			//SetDisabledButton(document.all.BtnCopy,false);
		}
		else
		{
			SetDisabledButton(document.all.BtnSave,false);
			//SetDisabledButton(document.all.BtnCopy,false);
		}
		SetDisabledButton(document.all.BtnSelect,false);
		SetDisabledButton(document.all.BtnBack,false);
		SetDisabledButton(document.all.BtnNew,true);
		
		SetHiddenButton(document.all.BtnSave,false);
		//SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnNew,true);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnBack,false);

	}
	
}

function UpdateCurrentObjFromScreen()
{
	CoverageSearchObj.XREFID = document.frames("TabFrame").document.frames.GetXREFID();
}

function IsCurrentSelectionValid()
{
	var XREFID = document.frames("TabFrame").document.frames.GetXREFID();
	
	if (XREFID == "")
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
	CoverageSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowSearchTab = true then %>
	AddTab("Search",120, "CoverageSearch-f.asp?MODE=<%=MODE%>&SearchAHSID=<%= Request.QueryString("SearchAHSID") %>", 1);
<%	end if 
	if bShowDetailsTab = true then %>
	AddTab("Details", 120, "CoverageDetails-f.asp?MODE=<%=MODE%>&XREFID=<%=Request.QueryString("XREFID")%>&AHSID=<%=Request.QueryString("AHSID")%>",1);
<%	end if %>

<%	If (Request.QueryString("XREFID") = "NEW" Or Request.QueryString("XREFID") = "") And bShowSearchTab <> false Then %>
	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
<%	else %>	
	SetActiveTabViaGet("Details");
<%	end if%>

}
</script>

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
</SCRIPT>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnNew_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		ChangeTabURL("Details", "CoverageDetails-f.asp?MODE=<%=MODE%>&XREFID=NEW");
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
				CoverageSearchObj.Selected = true;
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
			CoverageSearchObj.Selected = false;
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
<object VIEWASTEXT data="..\Scriptlets\TabScriptlet.htm" id="TabsControl" type="text/x-scriptlet"></object>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="1000">
</iframe>

<table>
<%	if bShowNew = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnNew" ACCESSKEY="N" LANGUAGE=javascript onclick="return BtnNew_onclick()"><u>N</u>ew</button></td>
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
