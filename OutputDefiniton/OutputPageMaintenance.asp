<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
		
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
	bShowLaunch = true
	
	Select Case CONTAINERTYPE
		Case "FRAMEWORK"
			bShowSelect = false
			bShowClose = false
		Case "MODAL"
			bShowBack = false
			bShowLaunch = false
		Case "DIALOG"
	End Select		

	Select Case MODE
		Case "RO"
			bShowNew = false
			bShowCopy = false
			bShowSave = false
			bShowDelete = false
			bShowLaunch = false
		Case "RW"
	End Select		

	Select Case DETAILONLY
		Case "TRUE"
			bShowNew = false
			bShowBack= false
			bShowCopy = false
			bShowSelect = false
			bShowDelete = false
			bShowSearchTab = false
			bShowClose = true
		Case "FALSE"
	End Select		
	

	If HasAddPrivilege("FNSD_OUTPUT_PAGE",SECURITYPRIV) <> True Then
		bShowNew = false
		bShowCopy = false
	End If
	
	If HasDeletePrivilege("FNSD_OUTPUT_PAGE",SECURITYPRIV) <> True Then
		bShowDelete = false
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

function COutputPageSearchObj()
{
	this.OPID = "";
	this.OPIDName = "";
	this.OPIDCaption = "";
	this.OPIDInputType = "";	
	this.Selected = false;	
}
var OutputPageSearchObj = new COutputPageSearchObj();

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
		SetHiddenButton(document.all.BtnLaunch,true);
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
		SetHiddenButton(document.all.BtnLaunch,false);

	}
	
}

function UpdateCurrentObjFromScreen()
{
	OutputPageSearchObj.OPID = document.frames("TabFrame").document.frames.GetOPID();
	OutputPageSearchObj.OPIDName = document.frames("TabFrame").document.frames.GetOPIDName();
	//OutputPageSearchObj.OPIDCaption = document.frames("TabFrame").document.frames.GetOPIDCaption();
	//OutputPageSearchObj.OPIDInputType = document.frames("TabFrame").document.frames.GetOPIDInputType();
}

function IsCurrentSelectionValid()
{
	var OPID = document.frames("TabFrame").document.frames.GetOPID();
	
	if (OPID == "")
		return false;
	else
		return true;
}	

function OnTabFramesReady()
{
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);

<%  if CONTAINERTYPE	= "MODAL" then %>
	OutputPageSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowSearchTab = true then %>
	AddTab("Search",120, "OutputPageSearch-f.asp?MODE=<%=MODE%>", 1);
<%	end if 
	if bShowDetailsTab = true then %>
	AddTab("Details", 120, "OutputPageDetails-f.asp?ODID=<%=Request.QueryString("ODID")%>&MODE=<%=MODE%>&OPID=<%=Request.QueryString("OPID")%>&CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>",1);
<%	end if %>

<%	if bShowSearchTab = true then %>
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
		ChangeTabURL("Details", "OutputPageDetails-f.asp?ODID=<%=Request.QueryString("ODID")%>&MODE=<%=MODE%>&OPID=NEW&CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>");
		SetActiveTabViaGet("Details");
		ConfigureButtonsOnActivateTab("Details");
		ConfigureButtonsOnNew();
	}		
}
function BtnBack_onclick()
{
<% If Request.QueryString("ODID") <> "" Then %>
	self.location.href = "OutputPageMaintenance.asp?CONTAINERTYPE=MODAL"
<% Else %>
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		document.all.TabsControl.SetActiveTab("1");
		if (DoOnscriptletevent("OnTabClick", "1") == true)
			ConfigureButtonsOnActivateTab("Search");
	}			
<% End If %>
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
			UpdateCurrentObjFromScreen();
		}
	}
}

function BtnLaunch_onclick()
{
	if (document.frames("TabFrame").document.readyState == "complete") 
	{
		lret = document.frames("TabFrame").document.frames.ExeLaunch()
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
				OutputPageSearchObj.Selected = true;
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
			//OutputPageSearchObj.Selected = false;
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

Function Handles(Obj, Title)
If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
	lret = top.frames("TOP").SetHandle(Obj, Title)
End If
End Function

Sub Window_Onload

End Sub
</SCRIPT>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<OBJECT data=..\Scriptlets\TabScriptlet.htm id=TabsControl style="LEFT: 0px; TOP: 0px" 
	type=text/x-scriptlet VIEWASTEXT></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="1000">
</iframe>
</body>
<TABLE>
<%	if bShowNew = true then %>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnNew" ACCESSKEY="N" LANGUAGE=javascript onclick="return BtnNew_onclick()"><u>N</u>ew</button></td>
<%	end if
	if bShowCopy = true then %>
<!--<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCopy" ACCESSKEY="C" LANGUAGE=javascript onclick="return BtnCopy_onclick()">Make <u>C</u>opy</button></td>-->
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
<%	end if 
	if bShowDelete = true then %>
<td width=20></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnDelete" LANGUAGE=javascript onclick="return BtnDelete_onclick()">Delete</button></td>
<%	end if 
	if bShowLaunch = true then %>
<td width=20></td>
<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnLaunch" LANGUAGE=javascript onclick="return BtnLaunch_onclick()">Visual Editor</button></td>
<% End If %>
</table>
</html>
