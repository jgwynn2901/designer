<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"

	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	If HasModifyPrivilege("FNSD_DICTIONARY",SECURITYPRIV) <> True Then MODE = "RO"

	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	
	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack
	dim bShowSearchTab
	
	bShowNew = true
	bShowCopy = true
	bShowSave = true 
	bShowSelect = true
	bShowClose = true
	bShowBack = true
	bShowSearchTab = true
	bShowDelete = true
	
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
			bShowSelect = false
			bShowSearchTab = false
		Case "FALSE"
	End Select		


	If HasAddPrivilege("FNSD_DICTIONARY",SECURITYPRIV) <> True Then
		bShowNew = false
		bShowCopy = false
		bShowDelete = false
	End If
%>


<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Dictionary Maintenance</title>
<%	if CONTAINERTYPE = "MODAL" then %>
<STYLE TYPE="text/css">
HTML {width: 335pt; height: 320pt}
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

function CDictSearchObj()
{
	this.DictText = "";
}
var DictSearchObj = new CDictSearchObj();


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
		SetDisabledButton(document.all.BtnSave,false);
		SetDisabledButton(document.all.BtnCopy,true);
		SetDisabledButton(document.all.BtnBack,true);
		SetDisabledButton(document.all.BtnNew,false);
		SetDisabledButton(document.all.BtnSelect,false);
		SetDisabledButton(document.all.BtnDelete,false);

		SetHiddenButton(document.all.BtnSave,true);
		SetHiddenButton(document.all.BtnCopy,true);
		SetHiddenButton(document.all.BtnBack,true);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnDelete,true);

	}
	if (inTab == "Details")
	{
		SetDisabledButton(document.all.BtnSave,false);
		SetDisabledButton(document.all.BtnCopy,false);
		SetDisabledButton(document.all.BtnBack,false);
		SetDisabledButton(document.all.BtnNew,false);
		SetDisabledButton(document.all.BtnSelect,false);
		SetDisabledButton(document.all.BtnDelete,false);

		SetHiddenButton(document.all.BtnSave,false);
		SetHiddenButton(document.all.BtnCopy,false);
		SetHiddenButton(document.all.BtnBack,false);
		SetHiddenButton(document.all.BtnNew,false);
		SetHiddenButton(document.all.BtnSelect,false);
		SetHiddenButton(document.all.BtnDelete,false);

	}
}

function UpdateCurrentObjFromScreen()
{
	DictSearchObj.DictText = document.frames("TabFrame").document.frames.GetDictText();
}

function IsCurrentSelectionValid()
{
	var RID = document.frames("TabFrame").document.frames.GetDictText();
	
	if (RID == "")
		return false;
	else
		return true;
}	

function OnTabFramesReady()
{
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);

<%  if CONTAINERTYPE	= "MODAL" then %>
	DictSearchObj	= window.dialogArguments;
<%	end if %>

		
<%	if bShowSearchTab = true then %>
	AddTab("Search",120, "DictionarySearch-f.asp?MODE=<%=MODE%>", 1);
<%	End If %>
	AddTab("Details", 120, "DictionaryEditor-f.asp?MODE=<%=MODE%>&RID=<%=Request.QueryString("RID")%>",1);

<%	If (Request.QueryString("RID") = "NEW" Or Request.QueryString("RID") = "") And bShowSearchTab <> false Then %>
	SetActiveTabViaGet("Search");
	ConfigureButtonsOnActivateTab("Search");
	<%	Else %>	
	SetActiveTabViaGet("Details");
<%	End If %>	
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
		ChangeTabURL("Details", "DictionaryEditor-f.asp?MODE=<%=MODE%>&RID=NEW");
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
				DictSearchObj.Selected = true;
				window.close();
			}
		}
	}		
}
function BtnDelete_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (document.frames("TabFrame").document.frames.ExeDelete() == true)
		{
			ConfigureButtonsOnSave();
			//DriverSearchObj.Saved = true;
		}
	}
}
function BtnClose_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		if (CanActiveFrameUnloadNow() == true)
		{
			DictSearchObj.Selected = false;
			window.close();
		}
	}		
}	
</SCRIPT>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE="VBScript">
Sub Document_OnKeyDown()
	If document.frames("TabFrame").document.readyState = "complete" Then
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
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="1000">
</iframe>

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
If bShowDelete = true then %>
	<td width=20></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnDelete" LANGUAGE=javascript onclick="return BtnDelete_onclick()">Delete</button></td>
	<td width=20></td>
<% End If
	if bShowBack = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnBack" LANGUAGE=javascript onclick="return BtnBack_onclick()">Back</button></td>
<%	end if %>

</TABLE>
</body>
</html>
