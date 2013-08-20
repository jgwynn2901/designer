<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<%	
Response.Expires = 0 
'***************************************************************
'General purpose: Displays a tree of available maintenance functions 
' on the left panel
'
'$History: mainttree.asp $ 
'* 
'* *****************  Version 11  *****************
'* User: Sohail.iqbal Date: 2/26/10    Time: 11:34a
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Changes are introduced as per MMAI-0055
'* 
'* *****************  Version 10  *****************
'* User: Jenny.cheung Date: 9/05/08    Time: 2:18p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* 
'* *****************  Version 10  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 2:01p
'* Updated in $/FNS_DESIGNER/Source/Designer/Maint
'* Added MyGreeting Interface to Designer for Sedgwick Only
'* 
'* *****************  Version 9  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:37p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Added MyGreeting interface
'* 
'* *****************  Version 8  *****************
'* User: Jenny.cheung Date: 1/25/07    Time: 9:06a
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Moved the Department interface to Account Related and created a new
'* permission FNSD_DEPARTMENT based on Doug's recommondation.
'* 
'* *****************  Version 7  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 1:38p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Added Department Interface due to ESIS Project.  It allows User to
'* create Department record attached to the AHSID in PROD Designer. The
'* permission used is the same as for Branch.  The Department sub folder
'* is located under Policy subtree which is under Maitenance Menu
'* 
'* *****************  Version 8  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 11:38a
'* Updated in $/FNS_DESIGNER/Source/Designer/Maint
'* Added Department Interface to Maintenance Tree due to the ESIS project
'* request. The data entry can only be done in PROD environment similar to
'* the Branch data entry setup as this Department_Codes table only resides
'* in the PROD.  It is located under Policy subtree folder and the
'* permission is setup based on the Branch permission.
'* 
'* *****************  Version 7  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:59p
'* Updated in $/FNS_DESIGNER/Source/Designer/Maint
'* Hartford SRS: Added Mailbox related stuff
'* 
'* *****************  Version 5  *****************
'* User: Alex.shimberg Date: 4/16/06    Time: 9:23p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Hartford SRS:
'* Added Claim Key Office Code Types link to the maintenance screens
'* 
'* *****************  Version 5  *****************
'* User: Alex.shimberg Date: 4/12/06    Time: 10:05p
'* Updated in $/FNS_DESIGNER/Source/Designer/Maint
'* Removed stop
'* 
'* *****************  Version 4  *****************
'* User: Alex.shimberg Date: 4/12/06    Time: 10:04p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Removed stop
'* 
'* *****************  Version 4  *****************
'* User: Alex.shimberg Date: 4/12/06    Time: 8:53p
'* Updated in $/FNS_DESIGNER/Source/Designer/Maint
'* New Claim Class Assignment module: Search, Details etc.
'* 
'* *****************  Version 3  *****************
'* User: Alex.shimberg Date: 4/10/06    Time: 10:56p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/Maint
'* Added Claim Class Assignment link

%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>FNS Account Lookup Tree</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

<!--#include file="..\lib\Help.asp"-->

Sub window_onresize()
	TreeView1.style.posTop = 0
	TreeView1.style.posLeft = 0
	TreeView1.style.pixelWidth = document.body.clientWidth
	TreeView1.style.height = document.body.clientHeight - 12
End Sub

Sub Window_Onload
<%
	'	Account related
	const ACCOUNTS = 1
	const BRANCH = 2
	const BRANCH_ASSIG_TYPE = 3
	const MANAGED_CARE = 4
	const CLAIM = 5
	const CONTACT = 6
	const EMPLOYEE = 7
	const ESCALATION = 8
	const ROUTING = 9
	const SPECIFIC_DESTINATION=10
	const OWNER = 11
	const GREETING =12
	const FIELD_HELP_INET=13 
	const COVERAGE_CODE_XREF=14 
	const MAILBOX = 15
	const MAILBOX_ASSIG_TYPE = 16
	const CLAIMCLASS = 17
	const DEPARTMENT =18	
	
	'	Attribute related
	const ATTRIB = 1
	const LOOKUP = 2
	const RULES = 3
	const DICTIONARY = 4
	
	'	Billing
	const FEE = 1
	const FEE_TYPE = 2

	'	Call Flow related
	const ADDR_BOOK = 1
	const CALL_FLOW = 2
	const FRAMES = 3
	
	'	Policy related
	const AGENT = 1
	const CARRIER = 2
	const COVERAGE = 3
	const DRIVER = 4
	const POLICY = 5
	const VEHICLE = 6
	const iNETPOLICY = 7
	const TPA = 8
				
	'	Routing Plan related
	const IMAGE_PREV = 1
	const OUTPUT_DEF = 2
	const OUTPUT_OVER = 3
	const OUTPUT_PAGES = 4
	const ROUTING_PL = 5

	'	Security related
	const USERS = 1
	const GROUPS = 2
	const MYGREETINGS = 3
				
	'	Vendors related
	const VENDORS = 1
	const NETWORKS = 2

	dim aAccount(18) , aAttribute(4), aBilling(2), aCallFlow(3), aPolicy(8), aRouting(5), aSecurity(3)
	dim aVendors(2)
	
	aAccount(ACCOUNTS) = HasViewPrivilege("FNSD_ACCOUNT_HIERARCHY_STEP","")
	aAccount(BRANCH) = HasViewPrivilege("FNSD_BRANCH","")
	aAccount(BRANCH_ASSIG_TYPE) = HasViewPrivilege("FNSD_BRANCH_ASSIGNMENT","")
	aAccount(MANAGED_CARE) = HasViewPrivilege("FNSD_MC_BRANCH_ASSIGNMENT","")
	aAccount(CLAIM) = HasViewPrivilege("FNSD_CLAIM_ASSIGNMENT","")
	aAccount(CLAIMCLASS) = HasViewPrivilege("FNSD_CLAIMCLASS_ASSIGNMENT","")
	aAccount(CONTACT) = HasViewPrivilege("FNSD_CONTACT","")
	aAccount(EMPLOYEE) = HasViewPrivilege("FNSD_EMPLOYEE","")
	aAccount(ESCALATION) = HasViewPrivilege("FNSD_ESCALATION","")
	aAccount(ROUTING) = HasViewPrivilege("FNSD_ROUTING_ADDRESS_RULE","")
	aAccount(SPECIFIC_DESTINATION) = HasViewPrivilege("FNSD_SPECIFIC_DESTINATION","")
	aAccount(OWNER) = HasViewPrivilege("FNSD_OWNER","")
	aAccount(GREETING)= HasViewPrivilege("FNSD_GREETING","")
	aAccount(COVERAGE_CODE_XREF)= HasViewPrivilege("FNSD_COVERAGE_CODE_XREF","")
	aAccount(FIELD_HELP_INET)= HasViewPrivilege("FNSD_FIELD_HELP_INETINTERNAL","")
	aAccount(MAILBOX)= HasViewPrivilege("FNSD_MAILBOX","")
	aAccount(MAILBOX_ASSIG_TYPE)= HasViewPrivilege("FNSD_MAILBOX_ASSIG_TYPE","")
	aAccount(DEPARTMENT) = HasViewPrivilege("FNSD_DEPARTMENT","") 

	aAttribute(ATTRIB) = HasViewPrivilege("FNSD_ATTRIBUTE","") 
	aAttribute(LOOKUP) = HasViewPrivilege("FNSD_LOOKUP_TYPES","")
	aAttribute(RULES) = HasViewPrivilege("FNSD_RULES","") 
	aAttribute(DICTIONARY) = HasViewPrivilege("FNSD_DICTIONARY","") 

	aBilling(FEE) = HasViewPrivilege("FNSD_FEE","")
	aBilling(FEE_TYPE) = HasViewPrivilege("FNSD_FEE_TYPE","")
	
	aCallFlow(ADDR_BOOK) = HasViewPrivilege("FNSD_ADDRESS_BOOK","")
	aCallFlow(CALL_FLOW) = HasViewPrivilege("FNSD_CALLFLOW","")
	aCallFlow(FRAMES) = HasViewPrivilege("FNSD_FRAMES","")	
	
	aPolicy(AGENT) = HasViewPrivilege("FNSD_AGENT","") 
	aPolicy(CARRIER) = HasViewPrivilege("FNSD_CARRIER","")
	aPolicy(TPA) = HasViewPrivilege("FNSD_TPA","") 
	aPolicy(COVERAGE) = HasViewPrivilege("FNSD_COVERAGE","") 
	aPolicy(DRIVER) = HasViewPrivilege("FNSD_DRIVER","") 
	aPolicy(POLICY) = HasViewPrivilege("FNSD_POLICY","") 
	aPolicy(VEHICLE) = HasViewPrivilege("FNSD_VEHICLE","") 
	aPolicy(iNETPOLICY) = HasViewPrivilege("FNSD_INETPOLICY","") 
	

	aRouting(IMAGE_PREV) = HasViewPrivilege("FNSD_IMAGE_PREVIEW","") 
	aRouting(OUTPUT_DEF) = HasViewPrivilege("FNSD_OUTPUT_DEFINITION","") 
	aRouting(OUTPUT_OVER) = HasViewPrivilege("FNSD_OUTPUT_OVERFLOW","")
	aRouting(OUTPUT_PAGES) = HasViewPrivilege("FNSD_OUTPUT_PAGE","")
	aRouting(ROUTING_PL) = HasViewPrivilege("FNSD_ROUTING_PLAN","")

	aSecurity(USERS) = HasViewPrivilege("FNSD_USERS","") Or HasAutomaticSecurityPrivilege()
	aSecurity(GROUPS) = HasViewPrivilege("FNSD_GROUPS","") Or HasAutomaticSecurityPrivilege()
	aSecurity(MYGREETINGS) = HasViewPrivilege("FNSD_MYGREETINGS","") Or HasAutomaticSecurityPrivilege()
	 
	aVendors(VENDORS) = HasViewPrivilege("FNSD_VENDORS","") Or HasAutomaticSecurityPrivilege()
	aVendors(NETWORKS) = HasViewPrivilege("FNSD_NETWORKS","") Or HasAutomaticSecurityPrivilege()

	if showThisOption(aAccount) then%>
 		TreeView1.AddNode "",1  , "STEP=1060",  "ACCRelated", "Account Related", "FOLDER", "FOLDERSEL" 
 		<%if aAccount(ACCOUNTS) then%>
 			TreeView1.AddNode "STEP=1060", 4 , "STEP=ACCT40", "ACCOUNT","Accounts", "PAGE", "PAGESEL" 	
 		<%end if
 		if aAccount(BRANCH) then%>
 			TreeView1.AddNode "STEP=1060", 4 , "STEP=10060", "BRANCH","Branch", "PAGE", "PAGESEL" 		
 		<%end if
 		if aAccount(BRANCH_ASSIG_TYPE) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=10070", "BRANCHASSIGN","Branch Assignment Types", "PAGE", "PAGESEL"  			
 		<%end if
 		if aAccount(CLAIM) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=10040", "CLAIMASSIGN","Claim Number Assignment Rules", "PAGE", "PAGESEL"  			
			TreeView1.AddNode "STEP=1060", 4 , "STEP=10090", "CLAIMKOC","Claim Key Office Code Types", "PAGE", "PAGESEL"  			
 		<%end if
 		if aAccount(CLAIMCLASS) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=100100", "CLAIMCLASS","Claim Class Assignment Rules", "PAGE", "PAGESEL"  			
		<%end if
 		if aAccount(CONTACT) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=CONT40", "CONTACT","Contact", "PAGE", "PAGESEL"  			
 		<%end if
 		if aAccount(EMPLOYEE) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=EMP40", "EMPLOYEE","Employee", "PAGE", "PAGESEL"  			
 		<%end if
 		if aAccount(ESCALATION) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=ESC240", "ESCALATION","Escalation Plan", "PAGE", "PAGESEL"  			
 		<%end if
 		if aAccount(MANAGED_CARE) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=MC10070", "MCBRANCHASSIGN","Managed Care Branch Assignment Types", "PAGE", "PAGESEL"  			
 		<%end if
 		if aAccount(ROUTING) then%>
			TreeView1.AddNode "STEP=1060", 4 , "STEP=10080", "ROUTINGADDRESS","Routing Address Rules", "PAGE", "PAGESEL"  			
 		<%end if
 	
 		if aAccount(SPECIFIC_DESTINATION) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10100", "SPECDEST","Specific Destinations", "PAGE", "PAGESEL"
		<%end if
		if aAccount(OWNER) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10120", "OWNER","Owner", "PAGE", "PAGESEL"
		<%end if
		
		if aAccount(GREETING) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10030", "GREETING","Greeting", "PAGE", "PAGESEL"
		<%end if
		if aAccount(FIELD_HELP_INET) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10020", "FIELDHELPINET","Field Help Inetinternal", "PAGE", "PAGESEL"
		<%end if
		if aAccount(COVERAGE_CODE_XREF) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10010", "COVERAGECODEXREF","Coverage Code XREF", "PAGE", "PAGESEL"
		<%end if
		if aAccount(MAILBOX) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10075", "MAILBOX","Mailbox", "PAGE", "PAGESEL"
		<%end if
		if aAccount(MAILBOX_ASSIG_TYPE) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10085", "MAILBOXASSIG","Mailbox Assignment Types", "PAGE", "PAGESEL"
		<%end if
		if aAccount(DEPARTMENT) then%>
 		 	TreeView1.AddNode "STEP=1060", 4 , "STEP=10086", "DEPARTMENT","Department", "PAGE", "PAGESEL"
		<%end if
	end if

	if showThisOption(aAttribute) then%>
		TreeView1.AddNode "",1  , "STEP=1030", "ARELATED","Attribute Related", "FOLDER", "FOLDERSEL" 
		<%if aAttribute(ATTRIB) then%>
			TreeView1.AddNode "STEP=1030", 4 , "STEP=1009", "ATTRIBUTE","Attributes", "PAGE", "PAGESEL" 
		<%end if
		if aAttribute(LOOKUP) then%>
			TreeView1.AddNode "STEP=1030", 4 , "STEP=1059", "MLUTYPES","Lookup Types", "PAGE", "PAGESEL"
		<%end if
		if aAttribute(RULES) then%>
			TreeView1.AddNode "STEP=1030", 4 , "STEP=1008", "MRULES","Rules", "PAGE", "PAGESEL" 
		<%end if
		if aAttribute(DICTIONARY) then%>
			TreeView1.AddNode "STEP=1030", 4 , "STEP=1007", "MDICTIONARY","Dictionary", "PAGE", "PAGESEL" 
		<%end if
	end if

	if showThisOption(aBilling) then%>
		TreeView1.AddNode "",1  , "STEP=AF1040", "BILLING","Billing", "FOLDER", "FOLDERSEL" 
		<%if aBilling(FEE) then%>
			TreeView1.AddNode "STEP=AF1040",4  , "STEP=XAF1040", "FEE","Fee", "PAGE", "PAGESEL" 
		<%end if
		if aBilling(FEE_TYPE) then%>
			TreeView1.AddNode "STEP=AF1040",4  , "STEP=XAAF1040", "FEE_TYPE","Fee Type", "PAGE", "PAGESEL" 
		<%end if
	end if		
		
	if showThisOption(aCallFlow) then%>
		TreeView1.AddNode "",1  , "STEP=1040", "CFRELATED","Call Flow Related", "FOLDER", "FOLDERSEL" 
		<%if aCallFlow(ADDR_BOOK) then%>
			TreeView1.AddNode "STEP=1040", 4 , "STEP=ADD40", "ADDRESS","Address Book", "PAGE", "PAGESEL" 
		<%end if
		if aCallFlow(CALL_FLOW) then%>
			TreeView1.AddNode "STEP=1040", 4 , "STEP=1099","CALLFLOW", "Call Flow", "PAGE", "PAGESEL" 
		<%end if
		if aCallFlow(FRAMES) then%>
			TreeView1.AddNode "STEP=1040", 4 , "STEP=1100","FRAMES", "Frames", "PAGE", "PAGESEL" 
		<%end if
	end if

	if showThisOption(aPolicy) then%>
		TreeView1.AddNode "",1  , "STEP=10160",  "PolicyRelated", "Policy Related", "FOLDER", "FOLDERSEL" 
		<%if aPolicy(AGENT) then%>
			TreeView1.AddNode "STEP=10160", 4 , "STEP=123660", "AGENT","Agent", "PAGE", "PAGESEL" 
		<%end if
		if aPolicy(CARRIER) then%>
			TreeView1.AddNode "STEP=10160", 4 , "STEP=100440", "CARRIER","Carrier", "PAGE", "PAGESEL" 
		<%end if
		if aPolicy(TPA) then%>
			TreeView1.AddNode "STEP=10160", 4 , "STEP=ZZ40", "TPA","TPA", "PAGE", "PAGESEL" 
		<%end if
		if aPolicy(COVERAGE) then%>
			'TreeView1.AddNode "STEP=10160", 4 , "STEP=ZZ40", "COVERAGE","Coverage", "PAGE", "PAGESEL" 
		<%end if	
		if aPolicy(DRIVER) then%>
			TreeView1.AddNode "STEP=10160", 4 , "STEP=A440", "DRIVER","Driver", "PAGE", "PAGESEL" 
		<%end if
		if aPolicy(POLICY) then%>
			TreeView1.AddNode "STEP=10160", 4 , "STEP=123330", "POLICY","Policy", "PAGE", "PAGESEL" 
		<%end if
		if aPolicy(VEHICLE) then%>
			'TreeView1.AddNode "STEP=10160", 4 , "STEP=123440", "VEHICLE","Vehicle", "PAGE", "PAGESEL" 
		<%end if
		if aPolicy(iNETPOLICY) then%>
			TreeView1.AddNode "STEP=10160", 4 , "STEP=123450", "INETPOLICY","iNet Policy", "PAGE", "PAGESEL"
		<%end if
	end if
	
	if showThisOption(aRouting) then%>
		TreeView1.AddNode "",1  , "STEP=1000",  "RPRelated", "Routing Plan Related", "FOLDER", "FOLDERSEL" 
		<%if aRouting(IMAGE_PREV) then%>
			TreeView1.AddNode "STEP=1000", 4 , "STEP=16642", "IMAGE","Image Preview", "PAGE", "PAGESEL" 
		<%end if
		if aRouting(OUTPUT_DEF) then%>
			TreeView1.AddNode "STEP=1000", 4 , "STEP=1002", "ODEFS","Output Definitions", "PAGE", "PAGESEL" 
		<%end if
		if aRouting(OUTPUT_OVER) then%>
			TreeView1.AddNode "STEP=1000", 4 , "STEP=10035", "OOVERFLOW","Output Overflow", "PAGE", "PAGESEL" 
		<%end if
		if aRouting(OUTPUT_PAGES) then%>
			TreeView1.AddNode "STEP=1000", 4 , "STEP=10045", "OPAGES","Output Pages", "PAGE", "PAGESEL" 
		<%end if
		if aRouting(ROUTING_PL) then%>
			TreeView1.AddNode "STEP=1000", 4 , "STEP=10054", "ROUTING","Routing Plans", "PAGE", "PAGESEL" 
		<%end if
	end if
	
	if showThisOption(aSecurity) then%>
		TreeView1.AddNode "",1  , "STEP=1050",  "SECRelated", "Security Related", "FOLDER", "FOLDERSEL" 
		<%if aSecurity(USERS) then%>
			TreeView1.AddNode "STEP=1050", 4 , "STEP=10050", "USERS","Users", "PAGE", "PAGESEL" 
		<%end if
		if aSecurity(GROUPS) then%>
			TreeView1.AddNode "STEP=1050", 4 , "STEP=10051", "GROUPS","Groups", "PAGE", "PAGESEL" 
		<%end if
		
		if aSecurity(MYGREETINGS) then%>
			TreeView1.AddNode "STEP=1050", 4 , "STEP=10052", "MYGREETINGS","My Greetings", "PAGE", "PAGESEL" 
		<%end if
	end if

	if showThisOption(aVendors) then%>
		TreeView1.AddNode "",1  , "STEP=1080", "VNDRELATED","Vendor Related", "FOLDER", "FOLDERSEL" 
		<%if aVendors(VENDORS) then%>
			TreeView1.AddNode "STEP=1080", 4 , "STEP=20040", "VENDORS","Vendors", "PAGE", "PAGESEL" 
		<%end if
		if aVendors(NETWORKS) then%>
			TreeView1.AddNode "STEP=1080", 4 , "STEP=20099","NETWORKS", "Networks", "PAGE", "PAGESEL" 
		<%end if
	end if%>
	call window_onresize()
End Sub

<%	
function showThisOption(aCheck)
dim nTop, x

showThisOption = false
nTop = UBound( aCheck )
for x = 1 to nTop	'	option base 1
	if aCheck( x ) then
		showThisOption = true
		exit for
	end if
next
end function	
%>
Sub TreeView1_NodeClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )

	Select Case NodeType
	
		Case "SPECDEST"
			Parent.frames("WORK").location.href = "../AH/SpecDestMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "ACCOUNT"
			Parent.frames("WORK").location.href = "../AH/AHSMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "ADDRESS"
			Parent.frames("WORK").location.href = "../AH/AddressBookMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "BRANCH"
			Parent.frames("WORK").location.href = "../Branch/BranchMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "BRANCHASSIGN"
			Parent.frames("WORK").location.href = "../BranchAssignment/BranchAssignTypeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "MCBRANCHASSIGN"
			Parent.frames("WORK").location.href = "../ManagedCareBranchAssignment/MCBranchAssignTypeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "CLAIMASSIGN"
			Parent.frames("WORK").location.href = "../ClaimAssignment/ClaimAssignRuleMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "CLAIMKOC"
			Parent.frames("WORK").location.href = "../ClaimKoc/ClaimKOCMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "CLAIMCLASS"
			Parent.frames("WORK").location.href = "../ClaimClass/ClaimClassMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "CONTACT"
		   Parent.frames("WORK").location.href = "../AH/ContactMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "OWNER"
			Parent.frames("WORK").location.href = "../AH/OwnerMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "COVERAGECODEXREF"
			Parent.frames("WORK").location.href = "../AH/CoverageMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "GREETING"
		     Parent.frames("WORK").location.href = "../AH/GreetingMaintenance.asp?CONTAINERTYPE=FRAMEWORK"	
		Case "FIELDHELPINET"
		     Parent.frames("WORK").location.href = "../AH/FieldHelpInetMaintenance.asp?CONTAINERTYPE=FRAMEWORK"	
	    Case "EMPLOYEE"
			Parent.frames("WORK").location.href = "../AH/EmployeeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "ESCALATION"
			Parent.frames("WORK").location.href = "../AH/EscalationMaintenance.asp?CONTAINERTYPE=FRAMEWORK"	
		Case "DEPARTMENT"
			Parent.frames("WORK").location.href = "../Department/DEPARTMENTMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		
								
		Case "ROUTINGADDRESS"
			Parent.frames("WORK").location.href = "../RoutingAddressRule/RoutingAddressRuleMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "ATTRIBUTE" 'Attributes
			Parent.frames("WORK").location.href = "../Attribute/AttributeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "MLUTYPES"
			Parent.frames("WORK").location.href = "../LookupType/LookupTypeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "MRULES" 'Rules
			Parent.frames("WORK").location.href = "../Rules/RuleMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "MDICTIONARY" 'Dictionary
			Parent.frames("WORK").location.href = "../Dictionary/DictionaryMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		
		
		Case "FEE"
			Parent.frames("WORK").location.href = "../Billing/BillingMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "FEE_TYPE"
			Parent.frames("WORK").location.href = "../Billing/FeeTypeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "CALLFLOW"
			Parent.frames("WORK").location.href = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=FRAMEWORK"
		Case "FRAMES"
			Parent.frames("WORK").location.href = "../CallFlow/FrameSearchModal.asp?CONTAINERTYPE=FRAMEWORK&FRAMEMAINT=Y"


		Case "AGENT"
			Parent.frames("WORK").location.href = "../Policy/AgentMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "CARRIER"
			Parent.frames("WORK").location.href = "../Policy/CarrierMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "TPA"
			Parent.frames("WORK").location.href = "../Policy/TPAMaintenance.asp?CONTAINERTYPE=FRAMEWORK"			
		Case "COVERAGE"
			Parent.frames("WORK").location.href = "../Policy/CoverageMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "DRIVER"		
			Parent.frames("WORK").location.href = "../Policy/DriverMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		'Case "OFFICE"
		'	Parent.frames("WORK").location.href = "../Policy/OfficeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "POLICY"
			Parent.frames("WORK").location.href = "../Policy/PolicyMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		'Case "PROPERTY"
		'	Parent.frames("WORK").location.href = "../Policy/PropertyMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "VEHICLE"
			Parent.frames("WORK").location.href = "../Policy/VehicleMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "INETPOLICY"
			Parent.frames("WORK").location.href = "../Policy/iNetPolicyMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		

		Case "ODEFS" ' Output Definitions
			Parent.frames("WORK").location.href = "../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "OOVERFLOW"
			Parent.frames("WORK").location.href = "../OutputDefiniton/OutputOverflowMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "OPAGES"
			Parent.frames("WORK").location.href = "../OutputDefiniton/OutputPageMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "ROUTING"
			Parent.frames("WORK").location.href = "../RoutingPlan/RoutingPlanSearchModal.asp?CONTAINERCONTEXT=DRILLIN&CONTAINERTYPE=FRAMEWORK"

		Case "USERS"
			Parent.frames("WORK").location.href = "../Users/UserMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "GROUPS"
			Parent.frames("WORK").location.href = "../Groups/GroupMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
			
		Case "MYGREETINGS"
			Parent.frames("WORK").location.href = "../MyGreetings/MyGreetingMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "IMAGE"
			Parent.frames("WORK").location.href = "../Images/BGBMP/Preview.asp"
		Case "VENDORS" 'Vendors
			Parent.frames("WORK").location.href = "../Vendors/VendorMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "NETWORKS" 'Networks
			Parent.frames("WORK").location.href = "../Vendors/NetworkMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "MAILBOX"
			Parent.frames("WORK").location.href = "../Mailbox/MailboxMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
		Case "MAILBOXASSIG"
			Parent.frames("WORK").location.href = "../MailboxAssignment/MailboxAssignTypeMaintenance.asp?CONTAINERTYPE=FRAMEWORK"
	End Select

End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case MenuItem

		Case "Node Search"
			showModalDialog  "SearchModal.asp?CFID=1"  , "PropertiesModal", "dialogWidth:700px;dialogHeight:500px"
		Case Else
	End Select
End Sub
-->
</script>
</head>
<body bgcolor="white" topmargin="0" leftmargin="0" RightMargin="0" bottommargin="0">
<table WIDTH="100%" Height="1" BGCOLOR="#006699" CELLPADDING="0" CELLSPACING="0">
<tr>
<td CLASS="LABEL" ALIGN="LEFT">
<font COLOR="WHITE">» Maintenance</td>
<td ALIGN="RIGHT" CLASS="LABEL"><img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8">&nbsp;</td>
</tr>
</table>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="100%">
<param NAME="ShowTips" VALUE="False">
</object>
</body>
</html>