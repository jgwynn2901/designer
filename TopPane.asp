<!--#include file="lib\common.inc"-->
<!--#include file="lib\security.inc"-->
<% Response.Expires = 0 
dim lHasAHSAccess, lHasMaintAccess, lHasMigrAccess, lHasRepAccess
dim lAHSBar, lMaintBar, lMigrBar, lRepBar
	
lAHSBar = false
lMaintBar = false
lMigrBar = false
lRepBar = false
lHasAHSAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_HIERARCHYTREE","")
lHasMaintAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_MAINTENANCE","")
lHasMigrAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_DATA_MIGRATION","")
lHasRepAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_REPORTS","")

if lHasAHSAccess then
	lAHSBar = true
end if
if lHasMaintAccess then
	lAHSBar = true
	lMaintBar = true
end if
if lHasMigrAccess then
	lAHSBar = true
	lMaintBar = true
	lMigrBar = true
end if
'if lHasRepAccess then
'	lAHSBar = true
'	lMaintBar = true
'	lMigrBar = true
'end if
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="FNSDESIGN.css">
</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 BGCOLOR=BLACK>
<script language="JavaScript" src="reports/toolbar.js"></script>
<script language="JavaScript">
function usertrackingObj() {
	this.ShowPolicyList = false;
	this.ShowCallFlowList = false;
	this.ShowRoutingPlanList = false;
	this.ShowUserList = false;
	this.ShowVendorList = false;
	this.ShowBillingList = false;
	this.ShowBranchList = false;
	this.ShowMCBranchList = false;
	this.ShowVendorEligibilityList = false;
	this.ShowFraudList = false;
	this.ShowSubrogationList = false;
	this.ShowMailbox = false;
	this.ShowDepartment = false;
	this.ShowTips = false;
}

function handleObj(inHandle, inTitle)
{
	this.handle = inHandle;
	this.title = inTitle;
}

var usertracking = new usertrackingObj();
var handleOutputDef = new handleObj(null,"Output Definition Editor");
var handleCallflowDef = new handleObj(null,"Callflow Definition Editor");
var handleArray = new Array(handleOutputDef, handleCallflowDef);

var ToolBar_Supported = ToolBar_Supported ;
var lIsFirst = true;
if (ToolBar_Supported != null && ToolBar_Supported == true)
{
	Frame_Supported = false;
	var environment = '<%=Session("ENVIRONMENT_ABBREVIATION")%>';    
	if  (environment== "P")  {
		setDefaultICPMenuColor("Red", "White", "#6495ed"); 
	}else if (environment == "PP")  {
		setDefaultICPMenuColor("Yellow", "Black", "#6495ed");	
	}else if(environment == "QA")  {
		setDefaultICPMenuColor("Blue", "White", "#6495ed"); 		
	}else if (environment == "BA"){ 
	setDefaultICPMenuColor("Black", "White", "#6495ed"); 
	}else { // Black as Default
	    setDefaultICPMenuColor("Black", "White", "#6495ed");
	}
	
	setToolbarBGColor("white");
	<%If lHasAHSAccess Then
		if not lAHSBar then%>
			lNoBar = true;
		<%else%>
			lNoBar = false;
		<%end if%>
		addICPMenu("HierarchyMenu", " Hierarchy", "Account Hierarchy Step Tree", "Designer-f.asp", "WORKAREA");
	<%end if
	If lHasMaintAccess Then
		if not lMaintBar then%>
			lNoBar = true;
		<%else%>
			lNoBar = false;
		<%end if%>
		addICPMenu("MaintenanceMenu", " Maintenance", "", "Maint/maint-f.asp", "WORKAREA");	
	<%end if
	If lHasMigrAccess Then
		if not lMigrBar then%>
			lNoBar = true;
		<%else%>
			lNoBar = false;
		<%end if%>
		addICPMenu("MigrationMenu", " Migration", "Data Migration", "Migration/Migration-f.asp", "WORKAREA");
	<%end if
	If lHasRepAccess Then%>
		lNoBar = false;	
		addICPMenu("ReportsMenu", "Reports =>", "Data Reports", "ReportMenu.asp", "_self");
	<%end if%>
	lNoBar = false;	
	addICPMenu("SettingsMenu", "Settings", "Save Current Settings", "", "window.showModalDialog(\"SettingsModal.asp\")", "WORKAREA");
	addICPMenu("HelpMenu", "Help", "Designer Help", "Help/Help-f.asp", "WORKAREA");
	lNoBar = true;
	addICPMenu("ExitMenu", "EXIT", "", "", "doExit()");
	drawToolbar();
}

function SetHandle(inHandle, inWhichHandle)
{
	if (inWhichHandle == "CALLFLOW") 
		handleCallflowDef.handle = inHandle;
	else if	(inWhichHandle == "OUTPUT") 
		handleOutputDef.handle = inHandle;
}

function doExit()
{
//var bContinue = CloseWindows();
//If (bContinue == true)
	window.setTimeout ("CheckSuccess()",1000, "JScript");
}

function CheckSuccess()
{
	var bSuccess = true;
	var idx;
	for (idx=0; idx < handleArray.length; idx++)
	{
		if (handleArray[idx].handle != null)
		{
			if (handleArray[idx].handle.closed == false)
			{
				alert("You must close the " + handleArray[idx].title + " window manually.");
				//check again for timing...
				if ((handleArray[idx].handle != null) && (handleArray[idx].handle.closed == false))
					handleArray[idx].handle.focus();
				bSuccess = false;	
				break;
			}
			else 
				handleArray[idx].handle = null;
		}
	}
	if (bSuccess == true) 
	{
		top.location.href = "login.asp"
		window.showModalDialog('logout.asp',0,'dialogcenter:yes;dialogwidth:20;dialogheight:10');
	}		
}	

function CloseWindows()
{
	for (i=0; i < handleArray.length; i++)
	{
		if ((handleArray[i].handle != null) && (handleArray[i].handle.closed == false))
		{
			bRes = confirm("The " + handleArray[i].title + " is open, would you like to close it?");
			if (bRes == true) 
				handleArray[i].handle.close();
			else 
			{
				handleArray[i].handle.focus();
				return false;
			}	
		}
	}
	return true;
}
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
<%
dim cConnName, cKey

cConnName= "Customized Connection"
For Each cKey in Application.Contents
	If Application.Contents(cKey) =Session("ConnectionString") Then
		cConnName = cKey
		Exit For
	End If
Next
%>
top.window.document.title = "FNSNet Designer (<%=cConnName%>)" 
parent.frames("WORKAREA").location.href = "blank.htm"
End Sub

-->
</SCRIPT>

</BODY>
</HTML>
