<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<%

	dim lCanAddToRoot, aNodes, x
	
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
    'Response.Write(Request.QueryString)

	MODE = "RW"
	DETAILONLY = "FALSE"
	SELECTONLY = "FALSE"
	CONTAINERTYPE = "MODAL"

	lCanAddToRoot = false
	aNodes = split( Session("ACCOUNT_SECURITY"), "," )
	for x=lbound(aNodes) to ubound(aNodes)
		if aNodes(x) = "1" then
			lCanAddToRoot = true
			exit for
		end if
	next
	if Request.QueryString("PARENT_AHSID") = "1" and not lCanAddToRoot then
		MODE = "RO"
	end if
    if Request.QueryString("OriginSource") = "USERS"  then
		ORIGIN = "USERS"
	end if

	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then 
		DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	end if
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then 
		CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	end if
	if CStr(Request.QueryString("SELECTONLY")) <> "" then 
		SELECTONLY = CStr(Request.QueryString("SELECTONLY"))
	end if
	If SELECTONLY <> "TRUE" Then
		If not HasModifyPrivilege("FNSD_ACCOUNT_HIERARCHY_STEP",SECURITYPRIV) Then 
			MODE = "RO"
		end if
	Else
		MODE ="RO"
	End If

	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack, bShowDelete
	dim bShowSearchTab, bShowDetailsTab
	bShowNew = true
	bShowCopy = true
	bShowSave = true 
	bShowSelect = true
	bShowClose = true
	bShowBack = true
	bShowDelete = true
	bShowSearchTab = true
	bShowDetailsTab = true

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
			bShowDelete = false
		Case "RW"
	End Select		

	Select Case DETAILONLY
		Case "TRUE"
			bShowNew = false
			bShowCopy = false
			bShowClose = false
			bShowSelect = false
			bShowDelete = false
			bShowSearchTab = false
		Case "FALSE"
	End Select		
	If not HasAddPrivilege("FNSD_ACCOUNT_HIERARCHY_STEP",SECURITYPRIV) Then
		bShowNew = false
		bShowCopy = false
	End If
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Search</title>
<%	if CONTAINERTYPE = "MODAL" then %>
<STYLE TYPE="text/css">
HTML {width: 500px; height: 510px}
</STYLE>
<%	end if %>
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

function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();

function ConfigureButtonsOnSave()
{
	SetDisabledButton(document.all.BtnCopy,false);
	SetDisabledButton(document.all.BtnNew,false);
	SetDisabledButton(document.all.BtnSelect,false);
	SetDisabledButton(document.all.BtnDelete,false);
}

function ConfigureButtonsOnDelete()
{
	SetDisabledButton(document.all.BtnSave,true);
	SetDisabledButton(document.all.BtnCopy,true);
	SetDisabledButton(document.all.BtnNew,false);
	SetDisabledButton(document.all.BtnSelect,true);
	SetDisabledButton(document.all.BtnDelete,true);
}

function ConfigureButtonsOnNew()
{
	SetDisabledButton(document.all.BtnSave,false);
	SetDisabledButton(document.all.BtnCopy, true);
	SetDisabledButton(document.all.BtnNew, true);
	SetDisabledButton(document.all.BtnSelect, true);
	SetDisabledButton(document.all.BtnDelete, true);
	
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
		SetDisabledButton(document.all.BtnDelete,true);

		SetHiddenButton(document.all.BtnSave,true);
		SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnBack,true);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnDelete,true);
	}
	else if (inTab == "Details")
	{
		if (IsCurrentSelectionValid()== false)
		{
			SetDisabledButton(document.all.BtnSave,true);
			SetDisabledButton(document.all.BtnCopy,true);
			SetDisabledButton(document.all.BtnDelete,true);
		}
		else
		{
			SetDisabledButton(document.all.BtnSave,false);
			SetDisabledButton(document.all.BtnCopy,false);
			SetDisabledButton(document.all.BtnDelete,false);
		}
		SetDisabledButton(document.all.BtnSelect,false);
		SetDisabledButton(document.all.BtnBack,false);
		SetDisabledButton(document.all.BtnNew,false);
		
		SetHiddenButton(document.all.BtnSave,false);
		SetHiddenButton(document.all.BtnCopy,false);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnBack,false);
		SetHiddenButton(document.all.BtnDelete,false);

	}
	
}

function UpdateCurrentObjFromScreen()
{
	AHSSearchObj.AHSID = document.frames("TabFrame").document.frames.GetAHSID();
	AHSSearchObj.AHSIDName = document.frames("TabFrame").document.frames.GetAHSIDName();
}

function IsCurrentSelectionValid()
{
	var AHSID = document.frames("TabFrame").document.frames.GetAHSID();
	
	if (AHSID == "")
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
	AHSSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowSearchTab = true then %>
	AddTab("Search",120, "AHSSearch-f.asp?MODE=<%=MODE%>&ORIGIN=<%=ORIGIN%>", 1);
<%	end if 
	if bShowDetailsTab = true then %>
	AddTab("Details", 120, "AHSDetails-f.asp?MODE=<%=MODE%>&AHSID=<%=Request.QueryString("AHSID")%>&PARENT_AHSID=<%= Request.QueryString("PARENT_AHSID") %>",1);
<%	end if %>

<%	If (Request.QueryString("AHSID") = "NEW" Or Request.QueryString("AHSID") = "") And bShowSearchTab = true Then %>
	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
<%	else %>	
	SetActiveTabViaGet("Details");
<%	end if%>

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
		ChangeTabURL("Details", "AHSDetails-f.asp?MODE=<%=MODE%>&AHSID=NEW&PARENT_AHSID=<%= Request.QueryString("PARENT_AHSID") %>");
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


function BtnDelete_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (document.frames("TabFrame").document.frames.ExeDelete() == true)
			ConfigureButtonsOnDelete();
	}			
}

function BtnSave_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (document.frames("TabFrame").document.frames.ExeSave() == true)
		{
			ConfigureButtonsOnSave();
			<% If Request.QueryString("PARENT_AHSID") <> "" Then %>
				Refresh()
			<% End If %>
			<% If Request.QueryString = "MODAL" Then %>
				UpdateCurrentObjFromScreen();
			<% End If %>
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
				AHSSearchObj.Selected = true;
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
			AHSSearchObj.Selected = false;
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

Sub BtnGrfxBack_OnClick()
dim cURL
dim lIsDirty
<%	
if MODE <> "RO" then
%>	
	lIsDirty = document.frames("TabFrame").document.frames.IsDirty()
	if lIsDirty then
		lIsDirty = msgbox("Data has changed. Leave page without saving?",vbYesNo,"FNSDesigner") = vbNo
	end if
	if not lIsDirty then
		cURL = "..\AH\NodeSummary.asp?DROPDOWN=POLICY&AHSID="
		<%
		If Request.QueryString("AHSID") = "NEW" Then
		%>
			cURL = cURL & "<%=Request.QueryString("PARENT_AHSID")%>"
		<%
		else
		%>
			cURL = cURL & "<%=Request.QueryString("AHSID")%>"
		<%
		End If
		%>
		location.href = cURL
	end if
<%
end if
%>
End Sub

Sub Refresh()
	window.setTimeout "top.frames(""WORKAREA"").document.frames(""LEFT"").location.reload()",1500
End Sub

</SCRIPT>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<% If Request.QueryString("CONTAINERTYPE") <> "MODAL" Then %>
<!--#include file="..\lib\NavBack.inc"-->
<% End If %>
<OBJECT data=..\Scriptlets\Tabscriptlet.htm id=TabsControl style="LEFT: 0px; TOP: 0px" 
	type=text/x-scriptlet VIEWASTEXT></OBJECT>
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
<%	end if  %>
</table>
</html>
