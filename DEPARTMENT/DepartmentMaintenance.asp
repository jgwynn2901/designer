<%
'***************************************************************
'Department search and details forms.
'
'$History: DepartmentMaintenance.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/25/07    Time: 9:09a
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/DEPARTMENT
'* Moved the Department interface to Account Related and created a new
'* permission FNSD_DEPARTMENT based on Doug's recommondation.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 1:39p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Policy
'* Added Department Interface due to ESIS Project.  It allows User to
'* create Department record attached to the AHSID in PROD Designer. The
'* permission used is the same as for Branch.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 12:10p
'* Created in $/FNS_DESIGNER/Source/Designer/Policy
'* Added Department Interface due to the ESIS Project.  It allows user to
'* attach AHSID to the department record.  Also, it allows user to delete,
'* create a new record and Edit an record in PROD Designer.  Permission
'* setup is the same as for Branch.  
'* 
'* *****************  Version 1  *****************
'* User: Jenny cheung Date: 01/19/2007    Time: 9:421am
'* Created in $/FNS_DESIGNER/Source/Designer/Policy
'* 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:45p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Policy
'* 
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	DETAILONLY = "FALSE"
	CONTAINERTYPE = "MODAL"
	SELECTONLY = "FALSE"
	RSAHSID = CStr(Request.QueryString("AHSID"))

	SECURITYPRIV = CStr(Request.QueryString("SECURITYPRIV"))
	if CStr(Request.QueryString("DETAILONLY")) <> "" then DETAILONLY = CStr(Request.QueryString("DETAILONLY"))
	if CStr(Request.QueryString("CONTAINERTYPE")) <> "" then CONTAINERTYPE = CStr(Request.QueryString("CONTAINERTYPE"))
	if CStr(Request.QueryString("SELECTONLY")) <> "" then SELECTONLY = CStr(Request.QueryString("SELECTONLY"))

	If SELECTONLY <> "TRUE" Then
		If HasModifyPrivilege("FNSD_BRANCH",SECURITYPRIV) <> True Then MODE = "RO"
	Else
		MODE ="RO"
	End If
		

		
	dim bShowNew, bShowCopy, bShowSave, bShowSelect, bShowClose, bShowBack
	dim bShowSearchTab, bShowDetailsTab
	bShowNew = true
	bShowCopy = false
	bShowSave = true 
	bShowSelect = true
	bShowClose = true
	bShowBack = false
	bShowSearchTab = true
	bShowDetailsTab = true

	Select Case CONTAINERTYPE
		Case "FRAMEWORK"
			bShowSelect = false
			bShowClose = false
		Case "MODAL"
			
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


	If HasAddPrivilege("FNSD_BRANCH",SECURITYPRIV) <> True Then
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

function CDepartmentSearchObj()
{
	this.DEPTID = "";
	this.Selected = false;	
}
var DepartmentSearchObj = new CDepartmentSearchObj();

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
	DepartmentSearchObj.DEPTID = document.frames("TabFrame").document.frames.GetDEPTID();
}

function IsCurrentSelectionValid()
{
	var DEPTID = document.frames("TabFrame").document.frames.GetDEPTID();
	
	if (DEPTID == "")
		return false;
	else
		return true;
}	

function OnTabFramesReady()
{
	document.all.TabFrame.style.height = document.body.clientHeight - (TabsControl.style.pixelHeight+30);

<%  if CONTAINERTYPE	= "MODAL" then %>
	DepartmentSearchObj = window.dialogArguments;
<%	end if %>
		
<%	if bShowSearchTab = true then %> 
	AddTab("Search",120, "DEPARTMENTSearch-f.asp?MODE=<%=MODE%>&SearchAHSID=<%= Request.QueryString("SearchAHSID") %>", 1);
<%	end if 
	if bShowDetailsTab = true then %>
	AddTab("Details", 120, "DEPARTMENTDetails-f.asp?MODE=<%=MODE%>&RSAHSID=<%=Request.QueryString("AHSID")%>&DEPTID=<%=Request.QueryString("DEPTID")%>",1);
<%	end if %>

<%	If (Request.QueryString("DEPTID") = "NEW" Or Request.QueryString("DEPTID") = "") And bShowSearchTab <> false Then %>
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
		ChangeTabURL("Details", "DEPARTMENTDetails-f.asp?MODE=<%=MODE%>&DEPTID=NEW");
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
				DepartmentSearchObj.Selected = true;
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
			DepartmentSearchObj.Selected = false;
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
