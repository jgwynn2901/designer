<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\ZIP.inc"-->

<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

	''Server.ScriptTimeout = 7200

	Dim AHSID, lHasClientNodeAccess, lHasParentNodeAccess, lHasPeerNodeAccess
	dim lResetMCAccess, lIsAgent, lIsSedgwich

	AHSID	= Request.QueryString("AHSID")

	lIsAgent = (Request.QueryString("PARENT_AHSID") = "23")

	if HasViewPrivilege("FNSD_CLIENT_NODE","") then
		lHasClientNodeAccess = ""
	else
		lHasClientNodeAccess = "disabled"
	end if
	if HasViewPrivilege("FNSD_PARENT_NODE","") then
		lHasParentNodeAccess = ""
	else
		lHasParentNodeAccess = "disabled"
	end if
	if HasViewPrivilege("FNSD_PEER_NODE","") then
		lHasPeerNodeAccess = ""
	else
		lHasPeerNodeAccess = "disabled"
	end if
	lResetMCAccess = HasViewPrivilege("FNSD_RESETMC","")
	AHSID	= Request.QueryString("AHSID")

'*********** MMAI 0019 Change ***********
Dim AHStype
If AHSID <> "" Then
	If AHSID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID = '" & AHSID & "'"
		Set RS = Conn.Execute(SQLST)

		if Not RS.EOF then
		 AHStype = RS("TYPE")
		end if
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if
end if
if AHStype = "ACCOUNT" then
	ACCOUNT_TYPE = "Account"
elseif AHStype = "INSURED" then
	ACCOUNT_TYPE = "Insured"
elseif AHStype = "RISK LOCATION" then
	ACCOUNT_TYPE = "Risk Location"
end if
'*********** MMAI 0019 Change ***********
If AHSID <> "" Then
	If AHSID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ACCOUNT_HIERARCHY_STEP AHS, AHS_VALID_RULES AVR,AHS_EXTENSION AHE WHERE AHS.ACCNT_HRCY_STEP_ID = " & AHSID & _
				" AND AHS.ACCNT_HRCY_STEP_ID = AVR.ACCNT_HRCY_STEP_ID(+) AND AHS.ACCNT_HRCY_STEP_ID = AHE.ACCNT_HRCY_STEP_ID (+)"
		Set RS = Conn.Execute(SQLST)

		if  Not RS.EOF then


			RSACCNT_HRCY_STEP_ID = RS("ACCNT_HRCY_STEP_ID")
			RSNODE_TYPE_ID = RS("NODE_TYPE_ID")
			RSPARENT_NODE_ID = RS("PARENT_NODE_ID")
			if not lIsAgent then
				if RSPARENT_NODE_ID = "23" then
					lIsAgent = true
				end if
			end if
			RSCLIENT_NODE_ID = RS("CLIENT_NODE_ID")

			lIsSedgwich = RSCLIENT_NODE_ID = "18"
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSPEER_NODE_ID = RS("PEER_NODE_ID")
			RSAUTO_ESCALATE = RS("AUTO_ESCALATE")
			RSTYPE = ReplaceQuotesInText(RS("TYPE"))
			RSFNS_CLIENT_CD = ReplaceQuotesInText(RS("FNS_CLIENT_CD"))
			RSADDRESS_1 = ReplaceQuotesInText(RS("ADDRESS_1"))
			RSADDRESS_2 = ReplaceQuotesInText(RS("ADDRESS_2"))
			RSADDRESS_3 = ReplaceQuotesInText(RS("ADDRESS_3"))
			RSCOUNTRY = ReplaceQuotesInText(RS("COUNTRY"))
			RSCITY = ReplaceQuotesInText(RS("CITY"))
			RSSTATE = RS("STATE")
			RSZIP = ReplaceQuotesInText(RS("ZIP"))
			RSFIPS = ReplaceQuotesInText(Trim(RS("FIPS")))
			RSCOUNTY = ReplaceQuotesInText(RS("COUNTY"))
			RSPHONE = ReplaceQuotesInText(RS("PHONE"))
			RSFAX = ReplaceQuotesInText(RS("FAX"))
			RSFEIN =ReplaceQuotesInText(RS("FEIN"))
			RSSIC = ReplaceQuotesInText(RS("SIC"))
			RSSUID = ReplaceQuotesInText(RS("SUID"))
			RSNATURE_OF_BUSINESS = ReplaceQuotesInText(RS("NATURE_OF_BUSINESS"))
			RSNODE_LEVEL = RS("NODE_LEVEL")
			RSLOCATION_CODE = ReplaceQuotesInText(RS("LOCATION_CODE"))
			RSLOCATION_NAME = ReplaceQuotesInText(RS("LOC_NAME"))
			RSESCALATION_CALLBACK_NUM = ReplaceQuotesInText(RS("ESCALATION_CALLBACK_NUM"))
			RSUPLOAD_KEY = ReplaceQuotesInText(RS("UPLOAD_KEY"))
			RSACTIVE_STATUS = RS("ACTIVE_STATUS")
			RSSTATUS_DATE = RS("STATUS_DATE")
			RSDEPT_NAME = ReplaceQuotesInText(RS("DEPT_NAME"))
			RSDEPT_CD = ReplaceQuotesInText(RS("DEPT_CD"))
			RSDIVISION_NAME = ReplaceQuotesInText(RS("DIVISION_NAME"))
			RSDIVISION_CD = ReplaceQuotesInText(RS("DIVISION_CD"))
			RSSEC_NAME = ReplaceQuotesInText(RS("SEC_NAME"))
			RSSEC_CD = ReplaceQuotesInText(RS("SEC_CD"))
			RSMANAGED_CARE_TYPE = ReplaceQuotesInText(RS("MANAGED_CARE_TYPE"))
			'RSPOLICY_SEARCH_ID = RS("POLICY_SEARCH_ID")
			RSCREATED_DT = RS("CREATED_DT")
			RSMODIFIED_DT = RS("MODIFIED_DT")
			RSADDITIONAL_DELIVERIES = RS("ADDITIONAL_DELIVERIES")
			SpecificDestination = RS.Fields("SPECIFIC_DESTINATION_FLAG")
			RSEMAIL = RS.Fields("EMAIL_ADDRESS")
			AGENT_BILLING_METHOD = RS.Fields("AGENT_BILLING_METHOD")
			AGENT_PAYMENT_TYPE = RS.Fields("AGENT_PAYMENT_TYPE")
			RSRULE_ID = RS.Fields("RULE_ID")
			if not isnull(RS.Fields("ACCOUNT_FROM_DATE")) then
				RSACCOUNT_FROM_DATE = cstr(cdate(RS.Fields("ACCOUNT_FROM_DATE")))
			end if
			if not isnull(RS.Fields("ACCOUNT_TO_DATE")) then
				RSACCOUNT_TO_DATE = cstr(cdate(RS.Fields("ACCOUNT_TO_DATE")))
			end if

			'MMAI-0007
			'Prashant Shekhar 05/21/2007
			'The requirement is to display ten fields as checkboxes in the Account Details frame
			'which maps to the AHS_Extension Table. Currently, the implementation covers
			'only ESIS.

		RS.Movefirst
		While  Not  RS.EOF

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:CONCENTRA_REVTELECLAIM_FLG" Then
				RSREV_TELECLAIM = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:SPECIAL_LOST_TIME_FLG" Then
				RSSPECIAL_LOST_TIME = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:SPECIAL_MED_ONLY_FLG" Then
				RSPECIAL_MED_ONLY = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:CONCENTRA_FIRST_SCRIPT_FLG" Then
				RSFIRST_SCRIPT = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:RN_TRIAGE_FLG" Then
				RSRN_TRIAGE = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:CONCENTRA_CAT_FCM_FLG" Then
				RSCAT_FCM = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:CONCENTRA_EXPO_PPO_FLG" Then
				RSEXPO_PPO = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:CONCENTRA_TCM_FLG" Then
				RSTCM = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:RO_INDICATOR" Then
				RSRO_INDICATOR = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:RO_ROUTING_FLG" Then
				RSRO_FLG = RS.Fields("FIELD_VALUE")
			end if
'MMAI-0023
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:OSHA_RECORDABLE" Then
				RSOSHA_RECORDABLE = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:LONGSHORE" Then
				RSLONGSHORE = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:GENERATE_EDI" Then
				RSEDI = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:CUSTOM_SEVERITY_FLG" Then
				RSSEVERITY = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:MONOPOLISTIC_STATE_HANDLIN" Then
				RSMONOPOLISTICSTATE = RS.Fields("FIELD_VALUE")
			end if

			'MMAI-0385
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:SELF_ADMIN_FLG" Then
				RSSELFADMININDICATOR = RS.Fields("FIELD_VALUE")
			end if

			'KKHU-0107
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:SRS_CLIENT" Then
				RSSRSCLIENT = RS.Fields("FIELD_VALUE")
			end if
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:HAS_LOCATION_NODE" Then
				RSHASLOCATIONNODE = RS.Fields("FIELD_VALUE")				
			end if
			'END-KKHU-0107

			'if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:FROM_DATE_TIP" Then
			'	RSACCOUNT_FROM_DATE_TIP = RS.Fields("FIELD_VALUE")
			'end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:FROM_DATE_TIP" OR RS.Fields("FIELD_NAME") = "CLAIM:INSURED:FROM_DATE_TIP" OR RS.Fields("FIELD_NAME") = "CLAIM:RISK_LOCATION:FROM_DATE_TIP"  Then
				RSACCOUNT_FROM_DATE_TIP = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:TO_DATE_TIP" OR RS.Fields("FIELD_NAME") = "CLAIM:INSURED:TO_DATE_TIP" OR RS.Fields("FIELD_NAME") = "CLAIM:RISK_LOCATION:TO_DATE_TIP" Then
				RSACCOUNT_TO_DATE_TIP = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:ACCOUNT_TYPE" Then
				RSACCOUNT_TYPE = RS.Fields("FIELD_VALUE")
			end if

			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:EMPLOYER_REPORT_LEVEL" Then
				RSEMPLOYER_REPORT_LEVEL = RS.Fields("FIELD_VALUE")
			end if
			
			'BCAB-0379
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:RO_OVERRIDE" Then
				RSRO_OVERRIDE = RS.Fields("FIELD_VALUE")
			end if
			
			'MMAI-0516
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:SECURE_EMAIL" Then
				RSSECURE_EMAIL = RS.Fields("FIELD_VALUE")
			end if
			
			'BCAB-0772
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:POLICY_LOOKUP_STATE" Then
				RSOLICY_LOOKUP_STATE = RS.Fields("FIELD_VALUE")
			end if
						
			'BCAB-0906
			if RS.Fields("FIELD_NAME") = "CLAIM:ACCOUNT:MASK_SSNO" Then
				RSMASK_SSN = RS.Fields("FIELD_VALUE")
			end if

			RS.Movenext
		Wend
		End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing
	Else
		RSPARENT_NODE_ID = Request.QueryString("PARENT_AHSID")

		If RSPARENT_NODE_ID <> "" And RSPARENT_NODE_ID <> "1" Then
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT CLIENT_NODE_ID FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID = " & RSPARENT_NODE_ID
			Set RS = Conn.Execute(SQLST)
			If Not RS.EOF then
				If IsNull(RS("CLIENT_NODE_ID")) Then
					RSCLIENT_NODE_ID = RSPARENT_NODE_ID
				Else
					RSCLIENT_NODE_ID = RS("CLIENT_NODE_ID")
				End If
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing
		end if
	end if


end if
%>
<html>
<head >
<meta name="VI60_defaultClientScript" content="VBScript">
<title>AHS Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CNodeSearchObj()
{
	this.AHSID = "";
	this.Selected = false;
}

function CAdditionalDeliveriesObj()
{
	this.strAdditional = "";
}


function SpecificDestObj()
{
	this.specDest = "";
}

var NodeSearchObj = new CNodeSearchObj();
var AdditionalDeliveriesObj = new CAdditionalDeliveriesObj();
var oSD = new SpecificDestObj();
var g_StatusInfoAvailable = false;

</script>
<!-- MMAI-0007
	 Prashant Shekhar 05/21/2007
	 he two javascripts below cover two requirements.The first one, DisplayCheckBox
	 displays the checkboxes only for ESIS. The second one, DefaultCheck_Insert covers
	 the requirement that while making a new insert some of the checkboxes are to checked
	 by default.
	 Currently, DisplayCheckBox() will be detached adn the checkboxes wil be displayed
	 for all cleints..   -->
<script id = "DisplayCheckBox" language = "javascript">

function DisplayCheckBox(AHS_ID)
{
	var hideTable;
	var clientnodeid;

	clientnodeid=document.FrmDetails.cOriginalClientNode.value//<%Request.QueryString("PARENT_AHSID")%>
	//alert(clientnodeid);
	hideTable = document.getElementById ("tblChkBox");

	//*********** MMAI 0019 Change ***********
	var type = document.FrmDetails.AHStype.value;
	//alert(type);
	if (isNaN(clientnodeid) || type != "ACCOUNT")
	{
		hideTable.style.display = "none";
	}
	//alert(AHS_ID);
	if (AHS_ID == "NEW")
	{
	hideTable.style.display = "block";
	}
	//*********** MMAI 0019 Change ***********
}

function DefaultCheck_Insert(AHS_ID)
{
	var clientnodeid;
	clientnodeid=document.FrmDetails.cOriginalClientNode.value
	if (clientnodeid == 202 || clientnodeid == 206)
	{
		if (AHS_ID == "NEW")
		{
			document.getElementById("Con_Cat_Loss").checked = "true";
			document.getElementById("Expo_Ind").checked = "true";
			document.getElementById("First_Script").checked = "true";
			document.getElementById("Rev_TeleClaim").checked = "true";
			document.getElementById("TCM_Ind").checked = "true";
			document.getElementById("Triage_Ind").checked = "true";
			//alert(AHS_ID);
		}
	}
	//********* SSAN-1273 ************
	if (AHS_ID == "NEW")
	{
		document.getElementById("CH_RO_OVERRIDE").checked = "true";
	}
	//********* SSAN-1273 ************
	//*********** MMAI 0019 Change ***********
	DisplayCheckBox(AHS_ID);
	//*********** MMAI 0019 Change ***********
}
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

<!--#include file="..\lib\Help.asp"-->

Sub window_onload
<%
if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%
end if
if AHSID <> "" then	%>
	document.all.NODE_TYPE_ID.value = "<%= RSNODE_TYPE_ID %>"
	document.all.STATE.value = "<%= RSSTATE %>"
	document.all.ACTIVE_STATUS.value = "<%= RSACTIVE_STATUS %>"
	document.all.MANAGED_CARE_TYPE.value = "<%= RSMANAGED_CARE_TYPE %>"
	document.all.TYPE.value = "<%= RSTYPE %>"
	<%
	If RSAUTO_ESCALATE = "Y" Then %>
		document.all.AUTO_ESCALATE.checked = true
	<%
	End If
	if lIsAgent then
		if AGENT_BILLING_METHOD = "" then
		%>
			document.all.AgentBillingNone.checked = true
		<%
		elseif AGENT_BILLING_METHOD = "M" then
		%>
			document.all.AgentBillingMonth.checked = true
		<%
		elseif AGENT_BILLING_METHOD = "Y" then
		%>
			document.all.AgentBillingYear.checked = true
		<%
		end if
		if AGENT_PAYMENT_TYPE = "" then
		%>
			document.all.AgentPayNone.checked = true
		<%
		elseif AGENT_PAYMENT_TYPE = "CHECK" then
		%>
			document.all.AgentPayCheck.checked = true
		<%
		elseif AGENT_PAYMENT_TYPE = "CREDIT" then
		%>
			document.all.AgentPayCC.checked = true
		<%
		end if
	end if
end if	'AHSID <> ""
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "AHSSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"
	FrmDetails.submit
End Sub

Sub UpdateAHSID(inAHSID)
	document.all.AHSID.value = inAHSID
	document.all.spanAHSID.innerText = inAHSID
End Sub

Sub UpdateClientNodeCD(inCD)
	document.all.CLIENT_NODE_ID.value = inCD
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub UpdateOriginalParent(inNewParentID)
	document.all.cOriginalParent.value = inNewParentID
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Function GetAHSID
	if document.all.AHSID.value <> "NEW" then
		GetAHSID = document.all.AHSID.value
	else
		GetAHSID = ""
	end if
End Function

Function GetAHSIDName
	GetAHSIDName = document.all.Name.value
End Function


Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	If  document.all.ACTIVE_STATUS.value = "" then
		MsgBox "Active Status is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	<%
	if lIsAgent then
	%>
		If (not document.all.AgentBillingYear.checked) AND (not document.all.AgentBillingMonth.checked) Then
			MsgBox "A 'Type of billing' Yearly or Monthly is required.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
		If (not document.all.AgentPayCheck.checked) AND (not document.all.AgentPayCC.checked) Then
			MsgBox "Payment method is a required field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	<%
	end if
	%>
	if len(document.all.ACC_FROM_DATE.value) <> 0 then
		if not isdate(document.all.ACC_FROM_DATE.value) then
			MsgBox "Please enter a valid date in field 'Account From Date'.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	end if
	if len(document.all.ACC_TO_DATE.value) <> 0 then
		if not isdate(document.all.ACC_TO_DATE.value) then
			MsgBox "Please enter a valid date in field 'Account To Date'.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	end if
	ValidateScreenData = true
End Function

sub UpdateScreenOnDelete()
	document.all.AHSID.value = ""
	FrmDetails.action = "AHSDetails.asp?STATUS=Delete successful."
	FrmDetails.target = "_self"
	FrmDetails.submit
end sub

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if

	if document.all.AHSID.value = "" then
		ExeCopy = false
		exit function
	end if
	document.all.AHSID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function Swap(ID)
If ID.checked = "True" Then
	Swap = "Y"
Else
	Swap = "N"
End If
End Function

function chkAgentBillType(cID)
If cID.AgentBillingYear.checked Then
	chkAgentBillType = "Y"
elseif cId.AgentBillingMonth.checked then
	chkAgentBillType = "M"
elseif cId.AgentBillingNone.checked then
	chkAgentBillType = ""
End If
End Function

function chkAgentPaymentType(cID)
If cID.AgentPayCheck.checked Then
	chkAgentPaymentType = "CHECK"
elseif cId.AgentPayCC.checked then
	chkAgentPaymentType = "CREDIT"
elseif cId.AgentPayNone.checked then
	chkAgentPaymentType = ""
End If
End Function

Function ExeSave
	sResult = ""
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if

	if document.all.AHSID.value = "" then
		ExeSave = false
		exit function
	end if

	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.AHSID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
			sResult = sResult & "CREATED_DT"& Chr(129) & "TO_DATE('" & date() & "', 'MM/DD/YY')" & Chr(129) & "0" & Chr(128)
			sResult = sResult & "STATUS_DATE"& Chr(129) & "TO_DATE('" & date() & "', 'MM/DD/YY')" & Chr(129) & "0" & Chr(128)
			sResult = sResult & "CURRENT_FLG"& Chr(129) & "'Y'" & Chr(129) & "0" & Chr(128)
		else
			document.all.TxtAction.value = "UPDATE"
			sResult = sResult & "MODIFIED_DT"& Chr(129) & "TO_DATE('" & date() & "', 'MM/DD/YY')" & Chr(129) & "0" & Chr(128)
		end if

		if ValidateScreenData = false then
			ExeSave = false
			exit function
		end if

		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID.value & Chr(129) & "1" & Chr(128)

		sResult = sResult & "NAME"& Chr(129) & document.all.NAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CLIENT_NODE_ID"& Chr(129) & document.all.CLIENT_NODE_ID.value & Chr(129) & "1" & Chr(128)

		sResult = sResult & "PARENT_NODE_ID"& Chr(129) & document.all.PARENT_NODE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PEER_NODE_ID"& Chr(129) & document.all.PEER_NODE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TYPE"& Chr(129) & document.all.TYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FNS_CLIENT_CD"& Chr(129) & document.all.FNS_CLIENT_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_1"& Chr(129) & document.all.ADDRESS_1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_2"& Chr(129) & document.all.ADDRESS_2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_3"& Chr(129) & document.all.ADDRESS_3.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.CITY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.STATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.ZIP.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FIPS"& Chr(129) & document.all.FIPS.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "COUNTRY"& Chr(129) & document.all.COUNTRY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "COUNTY"& Chr(129) & document.all.COUNTY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.PHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX"& Chr(129) & document.all.FAX.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FEIN"& Chr(129) & document.all.FEIN.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SIC"& Chr(129) & document.all.SIC.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SUID"& Chr(129) & document.all.SUID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NATURE_OF_BUSINESS"& Chr(129) & document.all.NATURE_OF_BUSINESS.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOCATION_CODE"& Chr(129) & document.all.LOCATION_CODE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOC_NAME"& Chr(129) & document.all.LOCATION_NAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "AUTO_ESCALATE"& Chr(129) & Swap(document.all.AUTO_ESCALATE) & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ESCALATION_CALLBACK_NUM"& Chr(129) & document.all.ESCALATION_CALLBACK_NUM.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NODE_TYPE_ID"& Chr(129) & document.all.NODE_TYPE_ID.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "NODE_LEVEL"& Chr(129) & document.all.NODE_LEVEL.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "UPLOAD_KEY"& Chr(129) & document.all.UPLOAD_KEY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACTIVE_STATUS"& Chr(129) & document.all.ACTIVE_STATUS.value & Chr(129) & "1" & Chr(128)

		sResult = sResult & "DEPT_NAME"& Chr(129) & document.all.DEPT_NAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DEPT_CD"& Chr(129) & document.all.DEPT_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DIVISION_NAME"& Chr(129) & document.all.DIVISION_NAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DIVISION_CD"& Chr(129) & document.all.DIVISION_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEC_NAME"& Chr(129) & document.all.SEC_NAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEC_CD"& Chr(129) & document.all.SEC_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MANAGED_CARE_TYPE"& Chr(129) & document.all.MANAGED_CARE_TYPE.value & Chr(129) & "1" & Chr(128)

		'sResult = sResult & "STATUS_DATE"& Chr(129) & document.all.STATUS_DATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDITIONAL_DELIVERIES" & Chr(129) & document.all.ADDITIONAL_DELIVERIES.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SPECIFIC_DESTINATION_FLAG" & Chr(129) & document.all.SpecificDestination.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "EMAIL_ADDRESS" & Chr(129) & document.all.EMailAddress.value & Chr(129) & "1" & Chr(128)

		if len(document.all.ACC_FROM_DATE.value) <> 0 then
			cFromDate = day(cdate(document.all.ACC_FROM_DATE.value)) & "-" & getMonth(month(cdate(document.all.ACC_FROM_DATE.value))) & "-" & year(cdate(document.all.ACC_FROM_DATE.value))
		else
			cFromDate = document.all.ACC_FROM_DATE.value
		end if
		sResult = sResult & "ACCOUNT_FROM_DATE" & Chr(129) & cFromDate & Chr(129) & "1" & Chr(128)
		if len(document.all.ACC_TO_DATE.value) <> 0 then
			cToDate = day(cdate(document.all.ACC_TO_DATE.value)) & "-" & getMonth(month(cdate(document.all.ACC_TO_DATE.value))) & "-" & year(cdate(document.all.ACC_TO_DATE.value))
		else
			cToDate = document.all.ACC_TO_DATE.value
		end if
		sResult = sResult & "ACCOUNT_TO_DATE" & Chr(129) & cToDate & Chr(129) & "1" & Chr(128)
		<%
		if lIsAgent then
		%>
		sResult = sResult & "AGENT_BILLING_METHOD" & Chr(129) & chkAgentBillType(document.all.AgentBilling) & Chr(129) & "1" & Chr(128)
		sResult = sResult & "AGENT_PAYMENT_TYPE" & Chr(129) & chkAgentPaymentType(document.all.AgentPayType) & Chr(129) & "1" & Chr(128)
		<%
		end if
		%>
		document.all.TxtSaveData.Value = sResult
		'*********** MMAI 0019 Change ***********
		document.all.CheckType.Value = document.all.TYPE.value
		'*********** MMAI 0019 Change ***********
		document.all.VALID_RULE_ID.value = document.all.RULE_ID.innerText
		document.all.FrmDetails.Submit()
		ClearDirty()
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If

	ExeSave = bRet

End Function

function getMonth(nMonth)
select case nMonth
	case 1
		getMonth = "JAN"
	case 2
		getMonth = "FEB"
	case 3
		getMonth = "MAR"
	case 4
		getMonth = "APR"
	case 5
		getMonth = "MAY"
	case 6
		getMonth = "JUN"
	case 7
		getMonth = "JUL"
	case 8
		getMonth = "AUG"
	case 9
		getMonth = "SEP"
	case 10
		getMonth = "OCT"
	case 11
		getMonth = "NOV"
	case 12
		getMonth = "DEC"
end select
end function

Sub TYPE_OnChange
	Dim s_OriginalType
	s_OriginalType = "<%= RSTYPE %>"
	IF document.all.TYPE.Value = "INSURED" AND document.all.MANAGED_CARE_TYPE.Value <> "" Then
		If MsgBox("At INSURED Level, Field (Managed Care Type) Is Not Applicable, the Field Will Get Blank Out." & VBCrLf & " Yes to Proceed, No to Cancel.", VBYesNo, "FNSDesigner") = VBYes Then
			document.all.MANAGED_CARE_TYPE.Value = ""
			document.all.MANAGED_CARE_TYPE.DISABLED = True
		Else 'VBNo
			document.all.TYPE.Value = s_OriginalType
			document.all.MANAGED_CARE_TYPE.DISABLED = False
		End If
	ELSEIF document.all.TYPE.Value = "INSURED" AND document.all.MANAGED_CARE_TYPE.Value = "" Then
		document.all.MANAGED_CARE_TYPE.Value = ""
		document.all.MANAGED_CARE_TYPE.DISABLED = True
	ELSE
		document.all.MANAGED_CARE_TYPE.DISABLED = False
	END IF
	Call Control_OnChange
End Sub

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"
	end if
end sub

sub StatusChange
<%
If AHSID <> "NEW" Then
%>
	msgbox "Please note that if you save this record," & vbcrlf & "the Status change will be applied to all the Nodes below this one as well.",0,"FNS Designer"
<%
end if
%>
	Control_OnChange
end sub

sub SetScreenFieldsReadOnly(bReadOnly, strNewClass)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnInput") = "TRUE" then
			document.all(iCount).readOnly = bReadOnly
			document.all(iCount).className = strNewClass
		elseif document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next
end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"
	End If
End Sub

Function EditDeliveries(ID)
	AdditionalDeliveriesObj.strAdditional = ID.value
	strURL = "..\AH\AdditionalDeliveriesModal.asp?MODE=<%=Request.QueryString("MODE")%>"
	showModalDialog strURL,AdditionalDeliveriesObj,"dialogWidth=430px; dialogHeight=270px; center=yes"

	ID.value = AdditionalDeliveriesObj.strAdditional
End Function

Sub EditSpecDest(cID)
If spanAHSID.innerText = "NEW" then
	msgbox "You must save this node before attaching a Specific Destination.",,"FNSDesigner"
else
	oSD.specDest = cID.value
	strURL = "AHSpecDestTop.asp?AHID=" & spanAHSID.innerText & "&NAME=" & f_EncodeURLString(document.all.NAME.Value) & "&SD=" & cID.value & "&CLIENT_NODE=" & document.all.CLIENT_NODE_ID.Value & "&MODE=<%=Request.QueryString("MODE")%>"
	window.showModalDialog strURL, oSD, "dialogWidth=570px; dialogHeight=420px; center=yes"
	cID.value = oSD.specDest
end if
End Sub

Function AttachRule (ID)
RID = ID.innerText
MODE = document.body.getAttribute("ScreenMode")
RuleSearchObj.RID = RID
'RuleSearchObj.RIDText = SPANID.innerhtml
RuleSearchObj.Selected = false

If RID = "" Then RID = "NEW"

	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID

showModalDialog  strURL  ,RuleSearchObj ,"dialogWidth:450px;dialogHeight:450px;center"
	SetDirty()
If RuleSearchObj.Selected = true Then
	If RuleSearchObj.RID <> ID.innerText then
		ID.innerText = RuleSearchObj.RID
	end if
	'SPANID.innerText = RuleSearchObj.RIDText
'ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
'	SPANID.innerText = RuleSearchObj.RIDText
End If

End Function

Function DetachRule(ID)
MODE = document.body.getAttribute("ScreenMode")
If MODE = "RO" Then
	Exit Function
End If
	SetDirty()
	ID.innerText = ""
	'SPANID.innerText = ""
End Function

Function AttachNode(ID)
	AHSID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	NodeSearchObj.AHSID = AHSID
	NodeSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"

	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No AHS currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If

	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ACCOUNT_HIERARCHY_STEP&SELECTONLY=TRUE&AHSID=" & AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"

	showModalDialog  strURL  ,NodeSearchObj ,"dialogWidth=650px; dialogHeight=700px; center=yes"
		If NodeSearchObj.AHSID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"
			ID.value = NodeSearchObj.AHSID
		end if
End Function

Function DetachNode(ID)

	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		exit Function
	end if

	ID.value = ""
	Setdirty()
End Function

Sub BtnResetMC_onclick
<%
If AHSID <> "NEW" then
%>
	if msgBox("This command will reset the MC Type of all the Nodes as well. Proceed?",4,"FNSDesigner") = 6 then
		document.all.TxtAction.value = "RESET"
		document.all.FrmDetails.Submit()
	end if
<%
end if
%>
End Sub

</script>
<script Language="JScript">
function f_EncodeURLString(in_SearchString){
	var s_OutPutString = new String();
	s_OutPutString = escape(in_SearchString);
	return s_OutPutString;
	}

function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}

var RuleSearchObj = new CRuleSearchObj();

</script>
</head>
<body onload = "javascript:DefaultCheck_Insert(document.all.AHSID.value);" topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10" valign="baseline"><nobr>&nbsp;» AHS Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8">&nbsp;</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form id="FrmDetails" Name="FrmDetails" METHOD="POST" ACTION="AHSSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchFNS_CLIENT_CD" value="<%=Request.QueryString("SearchFNS_CLIENT_CD")%>">
<input type="hidden" name="SearchNATURE_OF_BUSINESS" value="<%=Request.QueryString("SearchNATURE_OF_BUSINESS")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="AHSID" value="<%=Request.QueryString("AHSID")%>">
<input type="hidden" NAME="cOriginalStatus" value="<%=RSACTIVE_STATUS%>">
<input type="hidden" NAME="cOriginalParent" value="<%=RSPARENT_NODE_ID%>">
<input type="hidden" NAME="cOriginalValidRule" value="<%=RSRULE_ID%>" ID="Hidden2">
<input type="hidden" NAME="cOriginalClientNode" value="<%=RSCLIENT_NODE_ID%>">

<input type="hidden" NAME="VALID_RULE_ID" ID="Hidden1">

<input type="hidden" NAME="Search_TYPE" value="<%=Request.QueryString("Search_TYPE")%>">
<input type="hidden" NAME="SearchUPLOAD_KEY" value="<%=Request.QueryString("SearchUPLOAD_KEY")%>">
<input type="hidden" NAME="SearchLOCATION_CODE" value="<%=Request.QueryString("SearchLOCATION_CODE")%>">
<input type="hidden" NAME="SearchSUID" value="<%=Request.QueryString("SearchSUID")%>">
<!--*********** MMAI 0019 Change *********** -->
<input type="hidden" NAME="AHStype" value="<%=AHStype%>">
<input type="hidden" NAME="CheckType">
<!--*********** MMAI 0019 Change *********** -->
<% If AHSID <> "" OR AHSID = "NEW" Then %>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0" id="TblControls">
<tr>
<td>
<table class="LABEL">
<tr>
<td>A.H.S. ID:&nbsp;<span id="spanAHSID"><%=Request.QueryString("AHSID")%></span></td>
</tr>
</table>
<table cellspacing="0" cellpadding="0">
<tr>
<td CLASS="LABEL" WIDTH="70%" ALIGN="LEFT"><label STYLE="COLOR:black" CLASS="LABEL"><nobr>Created: <%= RSCREATED_DT %></label></td>
<td CLASS="LABEL" ALIGN="RIGHT"><label STYLE="COLOR:black" CLASS="LABEL"><nobr>Updated: <%= RSMODIFIED_DT %></label></td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL" COLSPAN="2">Name:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="NAME" CLASS="LABEL" SIZE="60" MAXLENGTH="80" VALUE="<%= RSNAME %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">FNS Client Code:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="FNS_CLIENT_CD" CLASS="LABEL" SIZE="3" MAXLENGTH="3" VALUE="<%= RSFNS_CLIENT_CD %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL">Address 1:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ADDRESS_1" CLASS="LABEL" SIZE="45" MAXLENGTH="45" VALUE="<%= RSADDRESS_1 %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">Phone:<br><input TYPE="TEXT" NAME="PHONE" ScrnInput="TRUE" CLASS="LABEL" SIZE="30" MAXLENGTH="20" VALUE="<%= RSPHONE %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Address 2:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ADDRESS_2" CLASS="LABEL" SIZE="45" MAXLENGTH="45" VALUE="<%= RSADDRESS_2 %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">Fax:<br><input TYPE="TEXT" NAME="FAX" ScrnInput="TRUE" CLASS="LABEL" SIZE="30" MAXLENGTH="20" VALUE="<%= RSFAX %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Address 3:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ADDRESS_3" CLASS="LABEL" SIZE="45" MAXLENGTH="45" VALUE="<%= RSADDRESS_3 %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">FEIN:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="FEIN" CLASS="LABEL" SIZE="30" MAXLENGTH="20" VALUE="<%= RSFEIN %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">SIC:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="SIC" CLASS="LABEL" SIZE="30" MAXLENGTH="6" VALUE="<%= RSSIC %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL" Colspan="2">Nature of Business:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="NATURE_OF_BUSINESS" CLASS="LABEL" SIZE="30" MAXLENGTH="30" VALUE="<%= RSNATURE_OF_BUSINESS %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>

<table>
<tr>
<td CLASS="LABEL">Zip:<br><input TYPE="TEXT" NAME="ZIP" ScrnInput="TRUE" CLASS="LABEL" SIZE="10" MAXLENGTH="20" VALUE="<%= RSZIP %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">City:<br><input TABINDEX="-1" TYPE="TEXT" NAME="CITY" CLASS="READONLY" readonly SIZE="30" MAXLENGTH="30" VALUE="<%= RSCITY %>"></td>
<td CLASS="LABEL">County:<br><input TYPE="TEXT" TABINDEX="-1" NAME="COUNTY" CLASS="READONLY" readonly SIZE="10" MAXLENGTH="30" VALUE="<%= RSCOUNTY %>"></td>
<td CLASS="LABEL">State:<br><input TYPE="TEXT" TABINDEX="-1" NAME="STATE" CLASS="READONLY" readonly SIZE="3" MAXLENGTH="3" VALUE="<%= RSSTATE %>"></td>
<td CLASS="LABEL">FIPS:<br><input TYPE="TEXT" NAME="FIPS" TABINDEX="-1" readonly CLASS="READONLY" SIZE="5" MAXLENGTH="5" VALUE="<%= Trim(RSFIPS) %>"><td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL">Country:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="COUNTRY" CLASS="LABEL" SIZE="22" MAXLENGTH="80" VALUE="<%= RSCOUNTRY %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">Node Type:<br>
<select NAME="NODE_TYPE_ID" CLASS="LABEL" STYLE="WIDTH:145" ScrnBtn="TRUE" ONCHANGE="VBScript::Control_OnChange">
<option VALUE>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
SQLNODE = ""
SQLNODE = SQLNODE & "SELECT * FROM NODE_TYPE ORDER BY NAME"
Set RS2 = Conn.Execute(SQLNODE)
Do WHile Not RS2.EOF
%>
<option VALUE="<%= RS2("NODE_TYPE_ID") %>"><%= RS2("NAME") %>
<%
RS2.MoveNext
Loop
RS2.Close
%>
</select>
</td>
<td CLASS="LABEL">Managed Care Type:<br>
<select NAME="MANAGED_CARE_TYPE" STYLE="WIDTH:100" ScrnBtn="TRUE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" <% If RSTYPE = "INSURED" Then Response.Write("DISABLED")%>>
<option VALUE>
<option VALUE="NEITHER">NEITHER
<option VALUE="CERTIFIED">CERTIFIED
<option VALUE="NOTCERTIFIED">NOTCERTIFIED
</select>
</td>
<tr>
</table>
<table>
<tr>
<td CLASS="LABEL" ALIGN="LEFT" WIDTH="10">SUID:<br><input TYPE="TEXT" NAME="SUID" ScrnInput="TRUE" CLASS="LABEL" SIZE="20" MAXLENGTH="20" VALUE="<%= RSSUID %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL">Escalation Call Back #:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ESCALATION_CALLBACK_NUM" CLASS="LABEL" SIZE="30" MAXLENGTH="10" VALUE="<%= RSESCALATION_CALLBACK_NUM %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL" VALIGN="BOTTOM"><input TYPE="CHECKBOX" ScrnBtn="TRUE" NAME="AUTO_ESCALATE" CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">Auto Escalate?</td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL">Type:<br>
<select NAME="TYPE" ID="TYPE" STYLE="WIDTH:168" ScrnBtn="TRUE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange">
<option VALUE>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
s_SQL = ""
s_SQL = s_SQL & "SELECT Value FROM VALID_VALUE WHERE Group_nm = 'ACCOUNT_TYPE' ORDER BY Value"
Conn.Open ConnectionString
Set rs_Type = Conn.Execute(s_SQL)
Do WHile Not rs_Type.EOF
%>
<option VALUE="<%= rs_Type("Value") %>"><%= rs_Type("Value") %>
<%
rs_Type.MoveNext
Loop
rs_Type.Close
%>
</select></td>


<td CLASS="LABEL">Active Status:<br>
<select NAME="ACTIVE_STATUS" STYLE="WIDTH:100" ScrnBtn="TRUE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::StatusChange">
<option VALUE="ACTIVE">Active
<option VALUE="COMBINED">Combined
<option VALUE="DEACTIVATED">Deactivated
</select>
</td>
<td>&nbsp;</td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL">Upload Key:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="UPLOAD_KEY" CLASS="LABEL" SIZE="80" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" MAXLENGTH="255" VALUE="<%= RSUPLOAD_KEY %>"></td>
</tr>
</table>
<table WIDTH="100%">
<tr>
<td CLASS="LABEL">Client Node ID:<br>
<img SRC="../Images/Attach.gif" <%=lHasClientNodeAccess%> ID="BtnATTACHCLIENTNODE" STYLE="CURSOR:HAND" ALT="Attach Client Node" OnClick="AttachNode(CLIENT_NODE_ID)" WIDTH="16" HEIGHT="16">
<img SRC="../Images/Detach.gif" <%=lHasClientNodeAccess%> ID="BtnDETACHCLIENTNODE" STYLE="CURSOR:HAND" ALT="Detach Client Node" OnClick="DetachNode(CLIENT_NODE_ID)" WIDTH="16" HEIGHT="16">
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="CLIENT_NODE_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="<%= RSCLIENT_NODE_ID %>"></td>
<td CLASS="LABEL" COLSPAN="2">Peer Node ID:<br>
<img SRC="../Images/Attach.gif" <%=lHasPeerNodeAccess%> ID="BtnATTACHPEERNODE" STYLE="CURSOR:HAND" ALT="Attach Peer Node" OnClick="AttachNode(PEER_NODE_ID)" WIDTH="16" HEIGHT="16">
<img SRC="../Images/Detach.gif" <%=lHasPeerNodeAccess%> ID="BtnDETACHPEERNODE" STYLE="CURSOR:HAND" ALT="Detach Peer Node" OnClick="DetachNode(PEER_NODE_ID)" WIDTH="16" HEIGHT="16">
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="PEER_NODE_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="<%= RSPEER_NODE_ID %>"></td>
<td CLASS="LABEL">Parent Node ID:<br>
<img SRC="../Images/Attach.gif" <%=lHasParentNodeAccess%> ID="BtnATTACHPARENTNODE" STYLE="CURSOR:HAND" ALT="Attach Parent Node" OnClick="AttachNode(PARENT_NODE_ID)" WIDTH="16" HEIGHT="16">
<img SRC="../Images/Detach.gif" <%=lHasParentNodeAccess%> ID="BtnDETACHPARENTNODE" STYLE="CURSOR:HAND" ALT="Detach Parent Node" OnClick="DetachNode(PARENT_NODE_ID)" WIDTH="16" HEIGHT="16">
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="PARENT_NODE_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="<%= RSPARENT_NODE_ID %>"></td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL">Location Name:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="LOCATION_NAME" CLASS="LABEL" SIZE="60" MAXLENGTH="80" VALUE="<%= RSLOCATION_NAME %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL" Colspan="2">Location Code:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="LOCATION_CODE" CLASS="LABEL" SIZE="15" MAXLENGTH="30" VALUE="<%= RSLOCATION_CODE %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Division Name:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="DIVISION_NAME" CLASS="LABEL" SIZE="60" MAXLENGTH="80" VALUE="<%= RSDIVISION_NAME %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL" Colspan="2">Division Code:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="DIVISION_CD" CLASS="LABEL" SIZE="15" MAXLENGTH="30" VALUE="<%= RSDIVISION_CD %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Department Name:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="DEPT_NAME" CLASS="LABEL" SIZE="60" MAXLENGTH="80" VALUE="<%= RSDEPT_NAME %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL" Colspan="2">Department Code:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="DEPT_CD" CLASS="LABEL" SIZE="15" MAXLENGTH="30" VALUE="<%= RSDEPT_CD %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Section Name:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="SEC_NAME" CLASS="LABEL" SIZE="60" MAXLENGTH="80" VALUE="<%= RSSEC_NAME %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td CLASS="LABEL" Colspan="2">Section Code:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="SEC_CD" CLASS="LABEL" SIZE="15" MAXLENGTH="30" VALUE="<%= RSSEC_CD %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL"><%=ACCOUNT_TYPE%> From Date:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ACC_FROM_DATE" CLASS="LABEL" SIZE="12" MAXLENGTH="12" VALUE="<%=RSACCOUNT_FROM_DATE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text1"></td>
<td CLASS="LABEL" Colspan="2"><%=ACCOUNT_TYPE%> To Date:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ACC_TO_DATE" CLASS="LABEL" SIZE="12" MAXLENGTH="12" VALUE="<%=RSACCOUNT_TO_DATE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text2"></td>
</tr>
<tr>
<td CLASS="LABEL"><%=ACCOUNT_TYPE%> From Date Tip:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ACC_FROM_DATE_TIP" CLASS="LABEL" SIZE="80" MAXLENGTH="255" VALUE="<%= RSACCOUNT_FROM_DATE_TIP %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="ACC_FROM_DATE_TIP"></td>
</tr>
<tr>
<td CLASS="LABEL"><%=ACCOUNT_TYPE%> To Date Tip:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="ACC_TO_DATE_TIP" CLASS="LABEL" SIZE="80" MAXLENGTH="255" VALUE="<%= RSACCOUNT_TO_DATE_TIP %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="ACC_TO_DATE_TIP"></td>
</tr>
<tr>
<td CLASS="LABEL">EMail Address:<br><input TYPE="TEXT" ScrnInput="TRUE" NAME="EMailAddress" CLASS="LABEL" SIZE="80" MAXLENGTH="80" VALUE="<%= RSEMAIL %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>

<table WIDTH="100%">
<tr>
<td CLASS="LABEL">Additional Deliveries:<br>
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="ADDITIONAL_DELIVERIES" CLASS="LABEL" SIZE="76" MAXLENGTH="255" VALUE="<%= RSADDITIONAL_DELIVERIES %>">
<img SRC="../Images/PropertiesIcon.gif" ID="BtnEDITDELIVERIES" STYLE="CURSOR:HAND" ALT="Edit Additional Deliveries" OnClick="EditDeliveries(ADDITIONAL_DELIVERIES)" WIDTH="16" HEIGHT="14"></td>
</tr>
</table>

<table WIDTH="100%">
<tr>
<td CLASS="LABEL">Specific Destination:<br>
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="SpecificDestination" CLASS="LABEL" SIZE="76" MAXLENGTH="255" VALUE="<%= SpecificDestination %>">
<img SRC="../Images/PropertiesIcon.gif" ID="BtnEDITSpecDest" STYLE="CURSOR:HAND" ALT="Edit Specific Destination(s)" OnClick="EditSpecDest(SpecificDestination)" WIDTH="16" HEIGHT="14"></td>
</tr>
</table>
<table class="Label" ID="Table1">
<tr>
<td CLASS="LABEL">Valid Rule:<br></td>
</tr>
<tr>
<td >
<img NAME="BtnAttachRule" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID">
<img NAME="BtnDetachRule" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule RULE_ID">
</td>
<td>Rule ID:&nbsp;<span ID="RULE_ID" CLASS="LABEL"><%=RSRULE_ID%></span></td>
</tr>
</table>
<!-- MMAI- 0007
	 Prashant Shekhar 05/21/2007
	 Display the checkboxe fields below the Valid Rule section in the Account Details frame. -->
<table ID = "tblChkBox" width="100%">
<tr>
<td>
	<table width="100%">
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "Con_Cat_Loss"  ScrnBtn="TRUE" NAME="CONCENTRA_CAT_LOSS" <%if RSCAT_FCM="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Concentra Cat Loss/FCM Referral Indicator</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "Expo_Ind" ScrnBtn="TRUE" NAME="EXPO_PPO_INDICATOR" <%if RSEXPO_PPO="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	EXPO/PPO Indicator</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "First_Script" ScrnBtn="TRUE" NAME="FIRST_SCRIPT_INDICATOR" <%if RSFIRST_SCRIPT="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	First Script Indicator</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "Rev_TeleClaim" ScrnBtn="TRUE" NAME="REVERSE_TELECLAIM_INDICATOR" <%if RSREV_TELECLAIM="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Reverse Teleclaim Indicator</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "TCM_Ind" ScrnBtn="TRUE" NAME="CONCENTRA_TCM_INDICATOR" <%if RSTCM="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Concentra TCM Indicator</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "Triage_Ind" ScrnBtn="TRUE" NAME="RN_TRIAGE_INDICATOR" <%if RSRN_TRIAGE="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	RN Triage Indicator</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "RO_Ind" ScrnBtn="TRUE" NAME="ACCOUNT_RECORD_INDICATOR"  <%if RSRO_INDICATOR="RO" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Account Record Only Indicator</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "RO_Rec" ScrnBtn="TRUE" NAME="GENERATE_ROUTING_RECORD"  <%if RSRO_FLG="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Generate Routing for Record Only</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "Lost_Time" ScrnBtn="TRUE" NAME="SPECIAL_LOST_TIME" <%if RSSPECIAL_LOST_TIME="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Special Lost Time Handling Indicator</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "Spec_Med" ScrnBtn="TRUE" NAME="SPECIAL_MEDICAL" <%if RSPECIAL_MED_ONLY="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Special Medical Only Handling Indicator</td>
	</tr>
<!--------------MMAI-0023-->
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "CH_OSHA_RECORDABLE" ScrnBtn="TRUE" NAME="CH_OSHA_RECORDABLE" <%if RSOSHA_RECORDABLE="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Osha Recordable</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "CH_LONGSHORE" ScrnBtn="TRUE" NAME="CH_LONGSHORE" <%if RSLONGSHORE="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Longshore, Federal, or Defense Based?</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "CH_EDI" ScrnBtn="TRUE" NAME="CH_EDI" <%if RSEDI="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Generate EDI/XML</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "CH_SEVERITY" ScrnBtn="TRUE" NAME="CH_SEVERITY" <%if RSSEVERITY="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Custom Severity Handling Indicator</td>
	</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">
	<input TYPE="CHECKBOX" ID = "CH_MONOPOLISTICSTATE" ScrnBtn="TRUE" NAME="CH_MONOPOLISTICSTATE" <%if RSMONOPOLISTICSTATE="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Monopolistic State Handling Indicator</td>
	<!------------------------- MMAI-0385 -->
	<td CLASS="LABEL" VALIGN="BOTTOM">
		<input TYPE="CHECKBOX" ID = "CH_SELFADMIN_INDICATOR" ScrnBtn="TRUE" NAME="CH_SELFADMIN_INDICATOR" <%if RSSELFADMININDICATOR="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Self Admin Indicator</td>
	</tr>
	<!---------------------------KKHU-0107-->
	<tr>
		<td CLASS="LABEL" VALIGN="BOTTOM">
		<input TYPE="CHECKBOX" ID = "CH_SRS_CLIENT" ScrnBtn="TRUE" NAME="CH_SRS_CLIENT" <%if RSSRSCLIENT="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		SRS Client</td>

		<td CLASS="LABEL" VALIGN="BOTTOM">
			<input TYPE="CHECKBOX" ID = "CH_HAS_LOCATION_NODE" ScrnBtn="TRUE"  NAME="CH_HAS_LOCATION_NODE" <%if RSHASLOCATIONNODE="Y" then Response.write "checked"%> CLASS="LABEL" DISABLED  >
		Has Location Node Record Type</td>
	</tr>

<!---------------------------BCAB-0379-->
	<tr>
		<td CLASS="LABEL" VALIGN="BOTTOM">
		<input TYPE="CHECKBOX" ID = "CH_RO_OVERRIDE" ScrnBtn="TRUE" NAME="CH_RO_OVERRIDE" <%if RSRO_OVERRIDE="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		Loss Severity Code Override</td>

		<!--MMAI-0516-->
		<td CLASS="LABEL" VALIGN="BOTTOM">
		<input TYPE="CHECKBOX" ID = "CH_SECURE_EMAIL" ScrnBtn="TRUE" NAME="CH_SECURE_EMAIL" <%if RSSECURE_EMAIL="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		Secure E-Mail</td>
	</tr>
	<tr>
		<td CLASS="LABEL" VALIGN="BOTTOM">
		<input TYPE="CHECKBOX" ID = "CH_MASK_SSN" ScrnBtn="TRUE" NAME="CH_MASK_SSN" <%if RSMASK_SSN="Y" then Response.write "checked"%> CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		Mask SSN</td>

	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">Account Type:
<select NAME="ACCOUNT_TYPE" STYLE="WIDTH:100" ScrnBtn="TRUE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
<option VALUE="O" <%if RSACCOUNT_TYPE ="O" then Response.write "selected"%>>OCIP
<option VALUE="C" <%if RSACCOUNT_TYPE ="C" then Response.write "selected"%>>CCIP
<option VALUE="N" <%if RSACCOUNT_TYPE ="N" then Response.write "selected"%>>Construction
<option VALUE="R" <%if RSACCOUNT_TYPE ="R" then Response.write "selected"%>>Regular
</select></td>
	<td CLASS="LABEL" VALIGN="BOTTOM">&nbsp;
	</td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">Employer Report Level:
<select NAME="SEL_EMPLOYER_REPORT_LEVEL" STYLE="WIDTH:200" ScrnBtn="TRUE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
<!--------------SEDGWICK ONLY -->
<%
if lIsSedgwich then
%>
<option VALUE="A" <%if RSEMPLOYER_REPORT_LEVEL="A" OR RSEMPLOYER_REPORT_LEVEL="" then Response.write "selected"%>>Account
<option VALUE="AE" <%if RSEMPLOYER_REPORT_LEVEL="AE" then Response.write "selected"%>>Account Entity
<option VALUE="AI" <%if RSEMPLOYER_REPORT_LEVEL="AI" then Response.write "selected"%>>Account - Insured
<option VALUE="AEI" <%if RSEMPLOYER_REPORT_LEVEL="AEI" then Response.write "selected"%>>Account Entity - Insured
<option VALUE="AR" <%if RSEMPLOYER_REPORT_LEVEL="AR" then Response.write "selected"%>>Account - Risk Location
<option VALUE="AER" <%if RSEMPLOYER_REPORT_LEVEL="AER" then Response.write "selected"%>>Account Entity - Risk Location
<option VALUE="I" <%if RSEMPLOYER_REPORT_LEVEL="I" then Response.write "selected"%>>Insured
<option VALUE="IE" <%if RSEMPLOYER_REPORT_LEVEL="IE" then Response.write "selected"%>>Insured Entity
<option VALUE="IA" <%if RSEMPLOYER_REPORT_LEVEL="IA" then Response.write "selected"%>>Insured - Account
<option VALUE="IEA" <%if RSEMPLOYER_REPORT_LEVEL="IEA" then Response.write "selected"%>>Insured Entity - Account
<option VALUE="IR" <%if RSEMPLOYER_REPORT_LEVEL="IR" then Response.write "selected"%>>Insured - Risk Location
<option VALUE="IER" <%if RSEMPLOYER_REPORT_LEVEL="IER" then Response.write "selected"%>>Insured Entity - Risk Location
<option VALUE="R" <%if RSEMPLOYER_REPORT_LEVEL="R" then Response.write "selected"%>>Risk Location
<option VALUE="RE" <%if RSEMPLOYER_REPORT_LEVEL="RE" then Response.write "selected"%>>Risk Location Entity
<option VALUE="RA" <%if RSEMPLOYER_REPORT_LEVEL="RA" then Response.write "selected"%>>Risk Location - Account
<option VALUE="REA" <%if RSEMPLOYER_REPORT_LEVEL="REA" then Response.write "selected"%>>Risk Location Entity - Account
<option VALUE="RI" <%if RSEMPLOYER_REPORT_LEVEL="RI" then Response.write "selected"%>>Risk Location - Insured
<option VALUE="REI" <%if RSEMPLOYER_REPORT_LEVEL="REI" then Response.write "selected"%>>Risk Location Entity - Insured
<%
else
%>
<option VALUE="A" <%if RSEMPLOYER_REPORT_LEVEL="A" OR RSEMPLOYER_REPORT_LEVEL="" then Response.write "selected"%>>Account
<option VALUE="I" <%if RSEMPLOYER_REPORT_LEVEL="I" then Response.write "selected"%>>Insured
<option VALUE="R" <%if RSEMPLOYER_REPORT_LEVEL="R" then Response.write "selected"%>>Risk Location
<option VALUE="L" <%if RSEMPLOYER_REPORT_LEVEL="L" then Response.write "selected"%>>Location Name
<%
end if
%>
</select>
	</td>
	<td CLASS="LABEL" VALIGN="BOTTOM">&nbsp;</td>
	</tr>
<!--------------END MMAI-0023-->
<!--BCAB-0772-->
	<tr>
		<td CLASS="LABEL" VALIGN="BOTTOM">Policy Lookup State:
			<select NAME="POLICY_LOOKUP_STATE" STYLE="WIDTH:200" ScrnBtn="TRUE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
				<option VALUE="RL" <%if RSOLICY_LOOKUP_STATE="RL" OR RSOLICY_LOOKUP_STATE="" then Response.write "selected"%>>Risk Location
				<option VALUE="LL" <%if RSOLICY_LOOKUP_STATE="LL" then Response.write "selected"%>>Loss Location
			</select>
		</td>
		<td CLASS="LABEL" VALIGN="BOTTOM">&nbsp;</td>
	</tr>
<!--End BCAB-0772-->
	</table>
</td>
</tr>
</table>


<%
if lIsAgent then
%>
<div>
<table ID="Table2">
<tr >
<td  CLASS="LABEL" width="114"><b><font size="1">For Agents only</font></b></td>
<td>
<hr>
</td>
</tr>
<tr>
<td>&nbsp</td>
<td>
<table border="0" bordercolor="#C0C0C0" cellspacing="1" width="301" ID="Table3">
<tr>
<td CLASS="LABEL" colspan="2" width="117">Type of billing:</td>
<td CLASS="LABEL" colspan="2" width="170">Payment method:</td>
</tr>
<tr>
<td width="7">&nbsp</td>
<td CLASS="LABEL" width="104" >
<input TYPE="radio" value="Year" NAME="AgentBilling" CLASS="LABEL" ID="AgentBillingYear">Yearly
</td>
<td CLASS="LABEL" width="338" >
<input TYPE="radio" value="Check" NAME="AgentPayType" CLASS="LABEL" ID="AgentPayCheck">Check
</td>
</TR>
<tr>
<td width="7">&nbsp</td>
<td CLASS="LABEL" width="104" >
<input TYPE="radio" value="Month" NAME="AgentBilling" CLASS="LABEL" ID="AgentBillingMonth">Monthly
</td>
<td CLASS="LABEL" width="338" >
<input TYPE="radio" value="Credit" NAME="AgentPayType" CLASS="LABEL" ID="AgentPayCC">Credit Card
</td>
</tr>
<tr>
<td width="7">&nbsp</td>
<td CLASS="LABEL" width="104" >
<input TYPE="radio" value="None" checked NAME="AgentBilling" CLASS="LABEL" ID="AgentBillingNone">None
</td>
<td CLASS="LABEL" width="104" >
<input TYPE="radio" value="None" checked NAME="AgentPayType" CLASS="LABEL" ID="AgentPayNone">None
</td>
</tr>
</table>
</td>
<td>
</td>
</tr>
<tr>
<td colspan="2" >
<hr>
<p>&nbsp;
</td>
</tr>
</table>
</div>
<%
end if
%>
<%
if lResetMCAccess then
%>
<table WIDTH="100%">
<tr><br></tr>
<tr>
<td CLASS="LABEL" width="35%">Reset Managed Care Type:</td>
<td><img SRC="../Images/reset.gif" ID="BtnResetMC" STYLE="CURSOR:HAND" ALT="Reset MC Types" width="62" height="26"></td>
</tr>
</table>
<%
end if
%>

<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
<%=Request.QueryString("STATUS") & "<br>"%>
No AHS selected.
</div>
<% End If %>
</form>
</body>
</html>


