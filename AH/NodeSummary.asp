<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<% 
dim lIsRoot, lIsCrawford

Response.Expires=0
Session("NodeID") = Request.QueryString("AHSID")
if Request.QueryString("AHSID") = "1" then
	lIsRoot = true
else
	lIsRoot = false
end if
If Request.QueryString("AHSID") <> "" Then 
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = SQL & "SELECT fns_client_cd, CLIENT_NODE_ID, NAME, ADDRESS_1, ADDRESS_2, ADDRESS_3, CITY, STATE, ZIP FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	Set RS = Conn.Execute(SQL)
	NAME = RS("NAME")
	ADDRESS_1 = RS("ADDRESS_1") 
	ADDRESS_2 = RS("ADDRESS_2") 
	ADDRESS_3 = RS("ADDRESS_3") 
	CITY = RS("CITY") 
	STATE = RS("STATE") 
	ZIP = RS("ZIP") 
	fns_client_cd = RS("fns_client_cd") 
	CLIENT_NODE_ID = RS("CLIENT_NODE_ID") 
End If
lIsCrawford = (CLIENT_NODE_ID = "81" or Request.QueryString("AHSID") = "81")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Business Entity Summary</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->

Dim PolicyToggle, RPToggle, CFToggle, VReferralToggle,VendorToggle, BillingToggle 
Dim UserToggle, BranchToggle, MCBranchToggle, FraudDetectionToggle, SubrogationToggle, MailboxToggle
Dim DepartmentToggle,TipsToggle

PolicyToggle="EXPAND"
RPToggle="EXPAND"
CFToggle="EXPAND"
VendorToggle="EXPAND"
BillingToggle="EXPAND"
UserToggle="EXPAND"
BranchToggle="EXPAND"
MCBranchToggle="EXPAND"
VReferralToggle="EXPAND"
FraudDetectionToggle="EXPAND"
SubrogationToggle="EXPAND"
MailboxToggle="EXPAND"
DepartmentToggle = "EXPLAND"
TipsToggle = "EXPLAND"

Sub BtnAgency_OnClick
	msgbox "Not Implemented Yet"
End Sub

Sub BtnBEDETAILS_OnClick
	Parent.frames("WORK").location.href = "NodeDetail.asp?AHID=1"
End Sub

Sub PolicyLoad_onclick
	LoadPolicy(PolicyToggle)
End Sub

sub MailboxLoad_OnClick
	LoadMailbox(MailboxToggle)
End Sub

sub DepartmentLoad_OnClick
	LoadDepartment(DepartmentToggle)
End Sub

Sub CallFlowLoad_onclick
	LoadCallFlow(CFToggle)
End Sub

Sub RoutingPlanLoad_onclick
	LoadRoutingPlan(RPToggle)
End Sub

Sub BillingLoad_onclick
	LoadBilling(BillingToggle)
End Sub

Sub UserLoad_OnClick
	LoadUser(UserToggle)
End Sub

Sub BillingLoad_onclick
	LoadBilling(BillingToggle)
End Sub

Sub PreferredVendorLoad_onclick
	LoadVendors(VendorToggle)
End Sub

Sub BranchLoad_OnClick
	LoadBranch(BranchToggle)
End Sub
Sub VendorReferralLoad_OnClick
	LoadVendorReferral(VReferralToggle)
End Sub
Sub MCBranchLoad_OnClick
	LoadMCBranch(MCBranchToggle)
End Sub

Sub VendorsLoad_OnClick
	LoadVendors(VendorToggle)
End Sub

Sub PolicyLoad_OnClick
	LoadPolicy(PolicyToggle)
End Sub

sub FraudDetecLoad_OnClick
	LoadFraudDetec(FraudDetectionToggle)
End Sub

sub SubrogationLoad_OnClick
	LoadSubrogation(SubrogationToggle)
End Sub

sub TipsLoad_OnClick
	LoadTips(TipsToggle)
End Sub


Sub LoadCallFlow(Action)
<%	If not HasViewPrivilege("FNSD_CALLFLOW","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	CFToggle=""
	CallFlowPlus.innerHTML = "- "
	document.all.CallFlowSummary.style.Pixelwidth = document.body.clientWidth
	document.all.CallFlowSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowCallFlowList = true
	document.all.CallFlowSummary.src = "AHCallFlowSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	CFToggle="EXPAND"
	CallFlowPlus.innerHTML = "+ "
	document.all.CallFlowSummary.style.Pixelwidth = "0"
	document.all.CallFlowSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowCallFlowList = false
End If
<%	End If %>
End Sub

Sub LoadBilling(Action)
<%	If not HasViewPrivilege("FNSD_FEE","") or lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	BillingToggle=""
	BillingPlus.innerHTML = "- "
	document.all.BillingSummary.style.Pixelwidth = document.body.clientWidth
	document.all.BillingSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowBillingList = true
	document.all.BillingSummary.src = "AHBillingSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	BillingToggle="EXPAND"
	BillingPlus.innerHTML = "+ "
	document.all.BillingSummary.style.Pixelwidth = "0"
	document.all.BillingSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowBillingList = false
End If
<%	End If%>
End Sub

Sub LoadPolicy(Action)
<%	If not HasViewPrivilege("FNSD_POLICY","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	PolicyToggle=""
	PolicyPlus.innerHTML = "- "
	document.all.PolicySummary.style.Pixelwidth = document.body.clientWidth
	document.all.PolicySummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowPolicyList = true
	document.all.PolicySummary.src = "AHPolicySummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	PolicyToggle = "EXPAND"
	PolicyPlus.innerHTML = "+ "
	document.all.PolicySummary.style.Pixelwidth = "0"
	document.all.PolicySummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowPolicyList = false
End If
<%	End If%>
End Sub

Sub LoadRoutingPlan(Action)
<%	If not HasViewPrivilege("FNSD_ROUTING_PLAN","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action="EXPAND" Then
	RoutingPlanPlus.innerHTML = "- "
	RPToggle=""
	document.all.RoutingPlanSummary.style.Pixelwidth = document.body.clientWidth
	document.all.RoutingPlanSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowRoutingPlanList = true
	document.all.RoutingPlanSummary.src = "AHRoutingPlanSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	RPToggle="EXPAND"
	RoutingPlanPlus.innerHTML = "+ "
	document.all.RoutingPlanSummary.style.Pixelwidth = "0"
	document.all.RoutingPlanSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowRoutingPlanList = false
End If
<%	End If%>
End Sub

Sub LoadUser(Action)
<%	If (not HasAutomaticSecurityPrivilege() And not HasViewPrivilege("FNSD_USERS","")) OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	UserToggle=""
	UserPlus.innerHTML = "- "
	document.all.UserSummary.style.Pixelwidth = document.body.clientWidth
	document.all.UserSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowUserList = true
	document.all.UserSummary.src = "AHAccountUserSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	UserToggle="EXPAND"
	UserPlus.innerHTML = "+ "
	document.all.UserSummary.style.Pixelwidth = "0"
	document.all.UserSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowUserList = false
End If
<%	End If%>
End Sub

Sub LoadBranch(Action)
<%	If not HasViewPrivilege("FNSD_BRANCH_ASSIGNMENT","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	BranchToggle=""
	BranchPlus.innerHTML = "- "
	document.all.BranchSummary.style.Pixelwidth = document.body.clientWidth
	document.all.BranchSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowBranchList = true
	document.all.BranchSummary.src = "AHBranchSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	BranchToggle="EXPAND"
	BranchPlus.innerHTML = "+ "
	document.all.BranchSummary.style.Pixelwidth = "0"
	document.all.BranchSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowBranchList = false
End If
<%	End If %>
End Sub

Sub LoadVendorReferral(Action)
<%	If not HasViewPrivilege("FNSD_COVERAGE_CODE_XREF","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	VReferralToggle=""
	VendorReferral.innerHTML = "- "
	document.all.VendorReferralSummary.style.Pixelwidth = document.body.clientWidth
	document.all.VendorReferralSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowVendorEligibilityList = true
	document.all.VendorReferralSummary.src = "AHVendorRefferalSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	VReferralToggle="EXPAND"
	VendorReferral.innerHTML = "+ "
	document.all.VendorReferralSummary.style.Pixelwidth = "0"
	document.all.VendorReferralSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowVendorEligibilityList = false
End If
<%	End If %>
End Sub

Sub LoadMCBranch(Action)
<%	If not HasViewPrivilege("FNSD_MC_BRANCH_ASSIGNMENT","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	MCBranchToggle=""
	MCBranchPlus.innerHTML = "- "
	document.all.MCBranchSummary.style.Pixelwidth = document.body.clientWidth
	document.all.MCBranchSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowMCBranchList = true
	document.all.MCBranchSummary.src = "AHMCBranchSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	MCBranchToggle="EXPAND"
	MCBranchPlus.innerHTML = "+ "
	document.all.MCBranchSummary.style.Pixelwidth = "0"
	document.all.MCBranchSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowMCBranchList = false
End If
<%	End If %>
End Sub

Sub LoadVendors(Action)
<%	If not HasViewPrivilege("FNSD_ACC_VENDOR","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	VendorToggle=""
	VendorsPlus.innerHTML = "- "
	document.all.VendorsSummary.style.Pixelwidth = document.body.clientWidth
	document.all.VendorsSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowVendorList = true
	document.all.VendorsSummary.src = "AccVendorSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	VendorToggle="EXPAND"
	VendorsPlus.innerHTML = "+ "
	document.all.VendorsSummary.style.Pixelwidth = "0"
	document.all.VendorsSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowVendorList = false
End If
<%	End If %>
End Sub

Sub LoadFraudDetec(Action)
<%	If not HasViewPrivilege("FNSD_FRAUD_DETECTION","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	FraudDetectionToggle=""
	FraudPlus.innerHTML = "- "
	document.all.FraudDetectionSummary.style.Pixelwidth = document.body.clientWidth
	document.all.FraudDetectionSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowFraudList = true
	document.all.FraudDetectionSummary.src = "FraudDetecSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	FraudDetectionToggle="EXPAND"
	FraudPlus.innerHTML = "+ "
	document.all.FraudDetectionSummary.style.Pixelwidth = "0"
	document.all.FraudDetectionSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowFraudList = false
End If
<%	End If %>
End Sub

Sub LoadSubrogation(Action)
<%	If not HasViewPrivilege("FNSD_SUBROGATION_DETECTION","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>

If Action = "EXPAND" Then
	SubrogationToggle=""
	SubrogationPlus.innerHTML = "- "
	document.all.SubrogationSummary.style.Pixelwidth = document.body.clientWidth
	document.all.SubrogationSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowSubrogationList = true
	document.all.SubrogationSummary.src = "SubrogationSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	SubrogationToggle="EXPAND"
	SubrogationPlus.innerHTML = "+ "
	document.all.SubrogationSummary.style.Pixelwidth = "0"
	document.all.SubrogationSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowSubrogationList = false
End If
<%	End If %>
End Sub

'AS 4/29/2006 - Added for Hartford SRS
Sub LoadMailbox(Action)
<%	If not HasViewPrivilege("FNSD_MAILBOX_ASSIGNMENT","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	MailboxToggle=""
	MailboxPlus.innerHTML = "- "
	document.all.MailboxSummary.style.Pixelwidth = document.body.clientWidth
	document.all.MailboxSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowMailbox = true
	document.all.MailboxSummary.src = "AHMailboxSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	MailboxToggle="EXPAND"
	MailboxPlus.innerHTML = "+ "
	document.all.MailboxSummary.style.Pixelwidth = "0"
	document.all.MailboxSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowMailbox = false
End If
<%	End If %>
End Sub

'AS 01/22/2007 - Added for ACE/ESIS
Sub LoadDepartment(Action)
<%	If not HasViewPrivilege("FNSD_BRANCH","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	DepartmentToggle=""
	DepartmentPlus.innerHTML = "- "
	document.all.DepartmentSummary.style.Pixelwidth = document.body.clientWidth
	document.all.DepartmentSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowDepartment = true
	document.all.DepartmentSummary.src = "AHLocationDeptSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	DepartmentToggle="EXPAND"
	DepartmentPlus.innerHTML = "+ "
	document.all.DepartmentSummary.style.Pixelwidth = "0"
	document.all.DepartmentSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowDepartment = false
End If
<%	End If %>
End Sub

'BCAB-0373 - Global Solution to Display AccountLOB Specific Tips
Sub LoadTips(Action)
<%	If not HasViewPrivilege("FNSD_TIP","") OR lIsRoot Then %>
	Exit Sub
<%	Else %>
If Action = "EXPAND" Then
	TipsToggle=""
	TipsPlus.innerHTML = "- "
	document.all.TipsSummary.style.Pixelwidth = document.body.clientWidth
	document.all.TipsSummary.style.Pixelheight = "185"
	top.frames("TOP").usertracking.ShowTips = true
	document.all.TipsSummary.src = "TipsSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
Else
	TipsToggle="EXPAND"
	TipsPlus.innerHTML = "+ "
	document.all.TipsSummary.style.Pixelwidth = "0"
	document.all.TipsSummary.style.Pixelheight = "0"
	top.frames("TOP").usertracking.ShowTips = false
End If
<%	End If %>
End Sub

Sub window_onload()
If top.frames("TOP").usertracking.ShowCallFlowList Then
	LoadCallFlow("EXPAND")
End If
If top.frames("TOP").usertracking.ShowRoutingPlanList Then
	LoadRoutingPlan("EXPAND")
End If
If top.frames("TOP").usertracking.ShowPolicyList Then	
	LoadPolicy("EXPAND")
End If	
If top.frames("TOP").usertracking.ShowBillingList Then	
	LoadBilling("EXPAND")
End If	
If top.frames("TOP").usertracking.ShowBranchList Then	
	LoadBranch("EXPAND")
End If
If top.frames("TOP").usertracking.ShowMCBranchList Then	
	LoadMCBranch("EXPAND")
End If
If top.frames("TOP").usertracking.ShowVendorEligibilityList Then	
	LoadVendorReferral("EXPAND")
End If
If top.frames("TOP").usertracking.ShowVendorList Then		
	LoadVendors("EXPAND")
End If
If top.frames("TOP").usertracking.ShowFraudList Then		
	LoadFraudDetec("EXPAND")
End If
If top.frames("TOP").usertracking.ShowSubrogationList Then		
	LoadSubrogation("EXPAND")
End If
If top.frames("TOP").usertracking.ShowUserList Then	
	LoadUser("EXPAND")
End If
If top.frames("TOP").usertracking.ShowMailbox Then	
	LoadMailbox("EXPAND")
End If
If top.frames("TOP").usertracking.ShowDepartment Then	
	LoadDepartment("EXPAND")
End If
If top.frames("TOP").usertracking.ShowTips Then	
	LoadTips("EXPAND")
End If
End Sub

Sub ExpandAll()
	LoadCallFlow("EXPAND")
	top.frames("TOP").usertracking.ShowCallFlowList = true
	LoadRoutingPlan("EXPAND")
	top.frames("TOP").usertracking.ShowRoutingPlanList = true
	LoadPolicy("EXPAND")
	top.frames("TOP").usertracking.ShowPolicyList = true
	LoadBilling("EXPAND")
	top.frames("TOP").usertracking.ShowBillingList = true
	LoadBranch("EXPAND")
	top.frames("TOP").usertracking.ShowBranchList = true
	LoadMCBranch("EXPAND")
	top.frames("TOP").usertracking.ShowMCBranchList = true
	LoadVendorReferral("EXPAND")
	top.frames("TOP").usertracking.ShowVendorEligibilityList = true
	LoadVendors("EXPAND")
	top.frames("TOP").usertracking.ShowVendorList = true
	LoadFraudDetec("EXPAND")
	top.frames("TOP").usertracking.ShowFraudList = true
	LoadSubrogation("EXPAND")
	top.frames("TOP").usertracking.ShowSubrogationList = true
	LoadUser("EXPAND")
	top.frames("TOP").usertracking.ShowUserList = true
	LoadMailbox("EXPAND")
	top.frames("TOP").usertracking.ShowMailBox = true
	LoadDepartment("EXPAND")
	top.frames("TOP").usertracking.ShowDepartment = true
	LoadTips("EXPAND")
	top.frames("TOP").usertracking.ShowTips = true
End Sub

Sub CollapseAll()
	LoadCallFlow("")
	top.frames("TOP").usertracking.ShowCallFlowList = false
	LoadRoutingPlan("")
	top.frames("TOP").usertracking.ShowRoutingPlanList = false
	LoadPolicy("")
	top.frames("TOP").usertracking.ShowPolicyList = false
	LoadBilling("")
	top.frames("TOP").usertracking.ShowBillingList = false
	LoadBranch("")
	top.frames("TOP").usertracking.ShowBranchList = false
	LoadMCBranch("")
	top.frames("TOP").usertracking.ShowMCBranchList = false
	LoadVendorReferral("")
	top.frames("TOP").usertracking.ShowVendorEligibilityList = false
	LoadVendors("")
	top.frames("TOP").usertracking.ShowVendorList = false
	LoadFraudDetec("")
	top.frames("TOP").usertracking.ShowFraudList = false
	LoadSubrogation("")
	top.frames("TOP").usertracking.ShowSubrogationList = false
	LoadUser("")
	top.frames("TOP").usertracking.ShowUserList = false
	LoadMailbox("")
	top.frames("TOP").usertracking.ShowMailbox = false
	LoadDepartment("")
	top.frames("TOP").usertracking.ShowDepartment = false
	LoadTips("")
	top.frames("TOP").usertracking.ShowTips = false
End Sub

Sub BtnAltName_onclick
	self.location.href = "AHS_ALT_Name.asp?AHSID=<%= Request.Querystring("AHSID") %>"
End Sub

Sub BtnContact_onclick
	self.location.href = "../Contacts/ContactDetailsData.asp?AHSID=<%= Request.Querystring("AHSID") %>"
End Sub

Sub BtnOwner_onclick
	self.location.href = "AHSOwners.asp?AHSID=<%= Request.Querystring("AHSID") %>"
End Sub

Sub BtnWCRPWizard_onclick
	self.location.href = "../routingplan/WCRPWizardinput.asp?AHSID=<%= Request.Querystring("AHSID")%>&CLIENT=<%=fns_client_cd%>"
End Sub

Sub BtnCRABTypes_onclick
	self.location.href = "CRA_BranchTy.asp?AHSID=<%= Request.Querystring("AHSID") %>"
End Sub

Sub BtnCRACovTypes_onclick
	self.location.href = "CRA_CoverTy.asp?AHSID=<%= Request.Querystring("AHSID") %>"
End Sub
-->
</script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
function BtnDetails_onclick() {
self.location.href = "AHSMaintenance.asp?AHSID=<%= Request.QueryString("AHSID") %>&DETAILONLY=TRUE"
}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" BGCOLOR="#d6cfbd" bottommargin="0" rightmargin="0">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Business Entity Summary&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Business Entity Summary.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table CELLPADDING="0" CELLSPACING="0" BORDER="0">
<tr><td>


<table CELLSPACING="0" CELLPADDING="0" WIDTH="300" BORDER="0" STYLE="BACKGROUND-COLOR:Seashell">
<tr>
<td CLASS="LABEL">
<b><nobr><%= NAME %></b> (<%= Request.QueryString("AHSID") %>)</td>
</tr>
<tr>
<td CLASS="LABEL"><nobr><%= ADDRESS_1 %></td>
</tr>
<td CLASS="LABEL"><nobr><%= ADDRESS_2 %></td>
<tr>
<td CLASS="LABEL"><nobr><%= ADDRESS_3 %></td>
</tr>
<tr>
<td CLASS="LABEL"><nobr><%= CITY %>
<% If CITY <> "" AND STATE <> "" Then 
Response.write(", ")
End If %><%= STATE %>&nbsp;<%= ZIP %></td>
</tr>
</table>

</td><td VALIGN="TOP">
&nbsp;
</td><td VALIGN="TOP">

<table CELLPADDING="0" CELLSPACING="15">
<tr>
<td ALIGN="CENTER" VALIGN="TOP">
<%if not lIsRoot then%>
<img SRC="../Images/accountdetails2.gif" TITLE="Account Details" ALT="Account Details" STYLE="CURSOR:HAND" ID="BtnDetails" LANGUAGE="javascript" onclick="return BtnDetails_onclick()" WIDTH="15" HEIGHT="14">
<%end if%>
</td>
<td ALIGN="CENTER" VALIGN="TOP">
<%if not lIsRoot then%>
<img SRC="../Images/alternatename2.gif" Title="Alternate Name" STYLE="CURSOR:HAND" ID="BtnAltName" WIDTH="15" HEIGHT="14">
<%end if%>
</td>
<td ALIGN="CENTER" VALIGN="TOP">
<%if 	not lIsRoot AND left(getInstanceName,3) <> "SED" then %>
<img SRC="../Images/contact.gif" Title="Contacts" STYLE="CURSOR:HAND" ID="BtnContact" WIDTH="15" HEIGHT="14">
<%end if%>
</td>
<td ALIGN="CENTER" VALIGN="TOP">
<%if not lIsRoot then%>
<img SRC="../Images/manager.gif" Title="Owners" STYLE="CURSOR:HAND" ID="BtnOwner" WIDTH="15" HEIGHT="14">
<%end if%>
</td>
<td ALIGN="CENTER" VALIGN="TOP">
<%if not lIsRoot then%>
<img SRC="../Images/bswoop.gif" Title="Workers Comp Routing Plan Wizard" STYLE="CURSOR:HAND" ID="BtnWCRPWizard" WIDTH="15" HEIGHT="14">
<%end if%>
</td>
<td ALIGN="CENTER" VALIGN="TOP">
<%if not lIsRoot then
	if lIsCrawford then%>
		<img SRC="../Images/branch.gif" Title="Crawford Branch Types" STYLE="CURSOR:HAND" ID="BtnCRABTypes" WIDTH="15" HEIGHT="14">
	<%end if
end if%>
</td>
<td ALIGN="CENTER" VALIGN="TOP">
<%if not lIsRoot then
	if lIsCrawford then%>
		<img SRC="../Images/MISC30.ICO" Title="Crawford Coverage Types" STYLE="CURSOR:HAND" ID="BtnCRACovTypes" WIDTH="15" HEIGHT="14">
	<%end if
end if%>
</td>
</tr>
</table>

</td><td VALIGN="TOP" WIDTH="100%" ALIGN="RIGHT">

<table BORDER="0">
<tr>
<%if not lIsRoot then%>
<td><img SRC="../IMAGES/ExpandAll.gif" STYLE STYLE="CURSOR:HAND" TITLE="Expand All" OnClick="ExpandAll()" WIDTH="13" HEIGHT="13"></td>
<td><img SRC="../IMAGES/CollapseAll.gif" STYLE STYLE="CURSOR:HAND" TITLE="Collapse All" OnClick="CollapseAll()" WIDTH="13" HEIGHT="13"></td>
<%end if%>
</tr>
</table>

</td></tr></table>

<%	If HasViewPrivilege("FNSD_CALLFLOW","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="3" HEIGHT="4"></td></tr>
<tr><td ID="CallFlowLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="CallFlowPlus">+ </span>Account Call Flows</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Account Call Flows.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table>
</td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="CallFlowSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<%	End If %>


<%	If HasViewPrivilege("FNSD_ROUTING_PLAN","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="RoutingPlanLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="RoutingPlanPlus">+ </span> Routing Plans</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="RoutingPlanSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<%	End If %>


<%	If HasViewPrivilege("FNSD_POLICY","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="PolicyLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="PolicyPlus">+ </span> Policy</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="PolicySummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<%	End If %>


<%	
		If (HasViewPrivilege("FNSD_USERS","") Or HasAutomaticSecurityPrivilege()) and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="UserLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="UserPlus">+ </span> Account Users</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="UserSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<%End If %>


<% If HasViewPrivilege("FNSD_FEE","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="BillingLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="BillingPlus">+ </span> Billing</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="BillingSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>


<% If HasViewPrivilege("FNSD_BRANCH_ASSIGNMENT","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="BranchLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="BranchPlus">+ </span> Branch Assignment Types</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="BranchSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

<% If HasViewPrivilege("FNSD_MC_BRANCH_ASSIGNMENT","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="MCBranchLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="MCBranchPlus">+ </span> Managed Care Branch Assignment Types</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="MCBranchSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

<% If HasViewPrivilege("FNSD_ACC_VENDOR","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="VendorReferralLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="VendorReferral">+ </span> Vendor Eligibility</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="VendorReferralSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

<% If HasViewPrivilege("FNSD_ACC_VENDOR","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="VendorsLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="VendorsPlus">+ </span> Vendor Selection</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="VendorsSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

<% If HasViewPrivilege("FNSD_FRAUD_DETECTION","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="FraudDetecLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="FraudPlus">+ </span> Fraud Detection</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="Table3">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="FraudDetectionSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>
<%
If HasViewPrivilege("FNSD_SUBROGATION_DETECTION","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table4">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="SubrogationLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="SubrogationPlus">+ </span> Subrogation</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table5">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="Table6">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="SubrogationSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

<% If HasViewPrivilege("FNSD_MAILBOX_ASSIGNMENT","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table7">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="MailboxLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="MailboxPlus">+ </span> Mailbox Assignment Types</td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table8">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="Table9">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="MailboxSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

<% If HasViewPrivilege("FNSD_DEPARTMENT","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table10">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="DepartmentLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="DepartmentPlus">+ </span> Department </td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table11">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="Table12">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="DepartmentSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>


<% If HasViewPrivilege("FNSD_TIP","") and not lIsRoot Then %>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table13">
<tr WIDTH="100%"><td colspan="2" HEIGHT="4"></td></tr>
<tr><td ID="TipsLoad" STYLE="CURSOR:HAND" CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;<span ID="TipsPlus">+ </span> Tips </td>
<td CLASS="GrpLabel">&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table14">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
</table>
<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="Table15">
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<iframe FRAMEBORDER="0" ID="TipsSummary" WIDTH="0" HEIGHT="0" SRC="ABOUT:BLANK">
</iframe>
<% End If %>

</body>
</html>

