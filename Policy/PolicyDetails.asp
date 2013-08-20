
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
Response.Expires=0
Response.Buffer = true

dim oConn, cSQL, oRS
Dim PID,CLIENTAHSID

NameTextLen = 30
PID	= CStr(Request.QueryString("PID"))
'MMAI 0007
'Prashant Shekhar 05/22/2007
CLIENTAHSID= CStr(Request.QueryString("AHSID"))



%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}

function CAgentSearchObj()
{
	this.AID = "";
	this.AIDName = "";
	this.Selected = false;
}
function CCarrierSearchObj()
{
	this.CID = "";
	this.CIDName = "";
	this.Selected = false;
}
function CTPASearchObj()
{
	this.TPAID = "";
	this.TPAIDName = "";
	this.Selected = false;
}

function CJurisdictionObj()
{
	this.Selected = false;
}

function CCoverageSearchObj()
{
	this.COVID = "";
	this.Selected = false;
	this.Saved = false;
}
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;
}

function CAdditionalDeliveriesObj()
{
	this.strAdditional = "";
}

var AHSSearchObj = new CAHSSearchObj();
var CoverageSearchObj = new CCoverageSearchObj();
var JurisdictionObj = new CJurisdictionObj();
var AgentSearchObj = new CAgentSearchObj();
var CarrierSearchObj = new CCarrierSearchObj();
var TPASearchObj = new CTPASearchObj();
var AdditionalDeliveriesObj = new CAdditionalDeliveriesObj();
var g_StatusInfoAvailable = false;


//function EditCoverage()
//{
	//if (InEditMode() == false) return;

	//PID = document.all.PID.value;
	//if ((PID == "") || (PID == "NEW")) return;

	//CVID = document.frames("CoverageFrame").GetSelectedCVID()
	//if (CVID != "")
	//{

		//alert("You do not have the appropriate security privileges to edit coverage.");
		//return;

		//CoverageSearchObj.Saved = false;
		//window.showModalDialog ("CoverageMaintenance.asp?SECURITYPRIV=FNSD_POLICY&CONTAINERTYPE=MODAL&DETAILONLY=TRUE&COVID=" + CVID, CoverageSearchObj, "center")
		//if (CoverageSearchObj.Saved == true)
			//RefreshCoverage();
	//}
	//else
		//alert ("Please choose a coverage to edit");
//}

//function NewCoverage()
//{
	//if (InEditMode() == false) return;

	//PID = document.all.PID.value;
	//if ((PID == "") || (PID == "NEW")) return;

//
		//alert("You do not have the appropriate security privileges to add coverage.");
		//return;
//

	//CoverageSearchObj.Saved = false;
	//window.showModalDialog ("CoverageMaintenance.asp?SECURITYPRIV=FNSD_POLICY&CONTAINERTYPE=MODAL&DETAILONLY=TRUE&COVID=NEW&PID=" + PID, CoverageSearchObj, "center")
	//if (CoverageSearchObj.Saved == true)
		//RefreshCoverage();
//}

//function RemoveCoverage()
//{
	//if (InEditMode() == false) return;

	//PID = document.all.PID.value;
	//if ((PID == "") || (PID == "NEW")) return;

	//CVID = document.frames("CoverageFrame").GetSelectedCVID()
	//if (CVID != "")
	//{
///
		//alert("You do not have the appropriate security privileges to delete coverage.");
		///return;

		//parent.frames("hiddenPage").location.href = "deletecoverage.asp?COVID=" + CVID
	//	RefreshCoverage();
	//}
	//else
		//alert ("Please choose a coverage to delete");

//}


//function RefreshCoverage()
//{
	//PID = document.all.PID.value;
	//document.frames("CoverageFrame").location.href = "PolicyDetailsCoverage.asp?PID=" + PID
//}

</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
if (document.all.CoverageFrame != null)
	document.all.CoverageFrame.style.height = .2 * document.body.clientHeight;
if (document.all.fldSet2 != null)
	document.all.fldSet2.style.height =  .2 * document.body.clientHeight;
if (document.all.JurisFrame != null)
	document.all.JurisFrame.style.height =  .2   *  document.body.clientHeight;
if (document.all.fldSet1 != null)
	document.all.fldSet1.style.height = .2 *   document.body.clientHeight;

</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmDetails.action = strURL
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"
	FrmDetails.submit
End Sub

Sub UpdatePID(inPID)
	document.all.PID.value = inPID
	document.all.spanPID.innerText = inPID
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Function GetPID
	GetPID = document.all.PID.value
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

Function checkDate( cInDate )
dim cTemp
'dim x
'dim aFunc(7,2)

checkDate = true
cTemp = trim( cInDate )
if len(cTemp) <> 10 then
	checkDate = false
elseif not IsNumeric(left(cTemp,2)) then
	checkDate = false
elseif not (cint(left(cTemp,2))>0 AND cint(left(cTemp,2))<13) then
	checkDate = false
elseif not ((asc(mid(cTemp,3,1)) = 45 AND asc(mid(cTemp,6,1)) = 45) OR (asc(mid(cTemp,3,1)) = 47 AND asc(mid(cTemp,6,1)) = 47)) then
	checkDate = false
elseif not IsNumeric(right(cTemp,4)) then
	checkDate = false
elseif not (cint(right(cTemp,4))>1990 AND cint(right(cTemp,4))<2100) then
	checkDate = false
elseif not isDate(cTemp) then
	checkDate = false
end if

'on error resume next
'checkDate = true
'aFunc(0,0) = "cTemp=rtrim(cstr(cInDate))"
'aFunc(0,1) = true			'	it's an assignment
'aFunc(1,0) = "len(cTemp)=10"
'aFunc(1,1) = false
'aFunc(2,0) = "IsNumeric(left(cTemp,2))"
'aFunc(2,1) = false
'aFunc(3,0) = "cint(left(cTemp,2))>0 AND cint(left(cTemp,2))<13"
'aFunc(3,1) = false
'aFunc(4,0) = "(asc(mid(cTemp,3,1)) = 45 AND asc(mid(cTemp,6,1)) = 45) OR (asc(mid(cTemp,3,1)) = 47 AND asc(mid(cTemp,6,1)) = 47)" 	'	- OR /
'aFunc(4,1) = false
'aFunc(5,0) = "IsNumeric(right(cTemp,4))"
'aFunc(5,1) = false
'aFunc(6,0) = "cint(right(cTemp,4))>1990 AND cint(right(cTemp,4))<2100"
'aFunc(6,1) = false
'aFunc(7,0) = "isDate(cTemp)"
'aFunc(7,1) = false
'x = 0
''	state machine
'do while checkDate and x <= ubound(aFunc)
'	if aFunc(x,1) then
'		execute aFunc(x,0)
'		if err.number <> 0 then
'			msgbox "Error executing " & aFunc(x,0)
'			msgbox err.description
'			msgbox "cInDate = " & cInDate
'			msgbox "We'll try executing as function..."
'			err.clear
'			eval(aFunc(x,0))
'			if err.number <> 0 then
'				msgbox "Still doesn't work!"
'			else
'				msgbox "It worked!!"
'			end if
'			checkDate = false
'			exit do
'		end if
'	else
'		checkDate = eval(aFunc(x,0))
'	end if
'	x = x + 1
'loop
end function

Function ValidateScreenData
errStr = ""
	If  document.all.TxtLOBCD.value = "" Then errStr = errStr & "LOB is a required field." & VBCRLF
	If  document.all.TxtMCTYPE.value = "" Then errStr = errStr & "Managed Care Type is a required field." & VBCRLF
	If  document.all.AHSID_ID.innerText = "" Then errStr = errStr & "A.H.Step ID is a required field." & VBCRLF
	If  document.all.TxtEffective.value = "" Then errStr = errStr & "Effective Date is a required field." & VBCRLF
	If  document.all.TxtExpiration.value = "" Then errStr = errStr & "Expiration Date  is a required field." & VBCRLF
	If  document.all.TxtLoad.value = "" Then errStr = errStr & "Load Date is a required field." & VBCRLF

	if errStr = "" then
		If Not CheckDate(document.all.TxtEffective.value) Then
			errstr = errstr & "Effective Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
		End If
		If Not CheckDate(document.all.TxtExpiration.value) Then
			errstr = errstr & "Expiration Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
		End If
		If document.all.TxtOriginalEffective.value <> "" then
			if Not CheckDate(document.all.TxtOriginalEffective.value) Then
				errstr = errstr & "Original Effective Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
			end if
		End If
		If document.all.TxtChange.value <> "" then
			if Not CheckDate(document.all.TxtChange.value) Then
				errstr = errstr & "Change Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
			end if
		End If
		If Not CheckDate(document.all.TxtLoad.value) Then
			errstr = errstr & "Load Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
		End If
		If document.all.TxtCancellation.value <> "" then
			if Not CheckDate(document.all.TxtCancellation.value) Then
				errstr = errstr & "Cancellation Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
			end if
		End If
	end if
	If errstr = "" Then
		ValidateScreenData = true
	Else
		MsgBox errstr, 0 , "FNSNetDesigner"
		ValidateScreenData = false
	End If
End Function

Function EditDeliveries(ID)
	AdditionalDeliveriesObj.strAdditional = ID.value
	strURL = "..\AH\AdditionalDeliveriesModal.asp?MODE="& "<%=Request.QueryString("MODE")%>"
	showModalDialog strURL,AdditionalDeliveriesObj,"center"

	ID.value = AdditionalDeliveriesObj.strAdditional
End Function

Function AttachAgent(ID,SPANID)
	AID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AgentSearchObj.AID = AID
	AgentSearchObj.AIDName = SPANID.innerText
	AgentSearchObj.Selected = false

	If AID = "" Then AID = "NEW"

	If AID= "NEW" And MODE = "RO" Then
		MsgBox "No agent currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If

	strURL = "AgentMaintenance.asp?SECURITYPRIV=FNSD_POLICY&CONTAINERTYPE=MODAL&AID=" & AID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"

	showModalDialog  strURL  ,AgentSearchObj ,"center"

	'if Selected=true update everything, otherwise if AID is the same, update text in case of save
	If AgentSearchObj.Selected = true Then
		If AgentSearchObj.AID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"
			ID.innerText = AgentSearchObj.AID
		end if
		UpdateFields SPANID,AgentSearchObj.AIDName
	ElseIf ID.innerText = AgentSearchObj.AID And AgentSearchObj.AID<> "" Then
		UpdateFields SPANID,AgentSearchObj.AIDName
	End If

End Function


Function Detach(ID,SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateFields(SPANID, fieldVal)
	If Len(fieldVal) < <%=NameTextLen%> Then
		SPANID.innertext = fieldVal
	Else
		SPANID.innertext = Mid ( fieldVal, 1, <%=NameTextLen%>) & " ..."
	End If
	SPANID.title = fieldVal
End Sub

Function AttachCarrier(ID,SPANID)
	CID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	CarrierSearchObj.CID = CID
	CarrierSearchObj.CIDName = SPANID.innerText
	CarrierSearchObj.Selected = false

	If CID = "" Then CID = "NEW"

	If CID= "NEW" And MODE = "RO" Then
		MsgBox "No carrier currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "CarrierMaintenance.asp?SECURITYPRIV=FNSD_POLICY&CONTAINERTYPE=MODAL&CID=" & CID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"

	showModalDialog  strURL  ,CarrierSearchObj ,"center"

	'if Selected=true update everything, otherwise if CID is the same, update text in case of save
	If CarrierSearchObj.Selected = true Then
		If CarrierSearchObj.CID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"
			ID.innerText = CarrierSearchObj.CID
		end if
		UpdateFields SPANID,CarrierSearchObj.CIDName
	ElseIf ID.innerText = CarrierSearchObj.CID And CarrierSearchObj.CID <> "" Then
		UpdateFields SPANID,CarrierSearchObj.CIDName
	End If

End Function

Function AttachTPA(ID,SPANID)
	TPAID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	TPASearchObj.TPAID = TPAID
	TPASearchObj.TPAIDName = SPANID.innerText
	TPASearchObj.Selected = false

	If TPAID = "" Then TPAID = "NEW"

	If TPAID= "NEW" And MODE = "RO" Then
		MsgBox "No TPA currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "TPAMaintenance.asp?SECURITYPRIV=FNSD_POLICY&CONTAINERTYPE=MODAL&TPAID=" & TPAID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"

	showModalDialog  strURL  ,TPASearchObj ,"center"

	'if Selected=true update everything, otherwise if CID is the same, update text in case of save
	If TPASearchObj.Selected = true Then
		If TPASearchObj.TPAID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"
			ID.innerText = TPASearchObj.TPAID
		end if
		UpdateFields SPANID,TPASearchObj.TPAIDName
	ElseIf ID.innerText = TPASearchObj.TPAID And TPASearchObj.TPAID <> "" Then
		UpdateFields SPANID,TPASearchObj.TPAIDName
	End If

End Function

Function AttachAccount (ID, SPANID)
	AHSID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID.title
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"

	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_POLICY&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"

	showModalDialog  strURL  ,AHSSearchObj ,"center"

	'if Selected=true update everything, otherwise if AHSID is the same, update text in case of save
	If AHSSearchObj.Selected = true Then
		If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"
			ID.innerText = AHSSearchObj.AHSID
		end if
		UpdateFields SPANID,AHSSearchObj.AHSIDName
	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		UpdateFields SPANID,AHSSearchObj.AHSIDName
	End If

End Function

Function GetSelectedCVID
	GetSelectedCVID = document.frames("CoverageFrame").GetSelectedCVID
End Function

Function GetSelectedState
	GetSelectedState = document.frames("JurisFrame").GetSelectedState
End Function

Sub ExeButtonsAttachJuris
	If Not InEditMode Then
		Exit Sub
	End If
	If document.all.PID.value = "" Or document.all.PID.value = "NEW" Then
		Exit Sub
	End If

	JurisdictionObj.Selected = false

	dim MODE
	MODE = document.body.getAttribute("ScreenMode")
	strURL = "PolicyJurisStateModal.asp?MODE=" & MODE & "&PID=" & document.all.PID.value
	showModalDialog  strURL,JurisdictionObj ,"center;scroll:no;resizable:yes"

	If JurisdictionObj.Selected = true Then	Refresh
End Sub

Sub Refresh
	PID = document.all.PID.value
	document.all.tags("IFRAME").item("JurisFrame").src = "PolicyDetailsJuris.asp?PID=" & PID
	'/document.all.tags("IFRAME").item("CoverageFrame").src = "PolicyDetailsCoverage.asp?PID=" & PID
End Sub

Function InEditMode
	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "This screen is read only.",0,"FNSNetDesigner"
		InEditMode = false
	End If
End Function

Function ExeCopy
	If Not InEditMode Then
		ExeCopy = false
		Exit Function
	End If

	If document.all.PID.value = "" Then
		ExeCopy = false
		Exit Function
	End If

	document.all.TxtAHSID.value = document.all.AHSID_ID.innerText

	FrmDetails.action = "PolicyCopy.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "hiddenPage"
	FrmDetails.submit
'	Refresh is done inside PolicyCopy.asp due to timing
	ExeCopy = true
End Function

Function ExeSave
    If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.PID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false

	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then

		if ValidateScreenData = false then
			ExeSave = false
			exit function
		end if

		If document.all.PID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "Policy_ID"& Chr(129) & document.all.PID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "POLICY_NUMBER"& Chr(129) & document.all.TxtNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "POLICY_DESC"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
		'MMAI-0023
		sResult = sResult & "POLICY_TYPE"& Chr(129) & document.all.selPolicyType.value & Chr(129) & "1" & Chr(128)
		'end
		sResult = sResult & "COMPANY_CODE"& Chr(129) & document.all.TxtCompanyCode.value & Chr(129) & "1" & Chr(128)

		sResult = sResult & "MANAGED_CARE_TYPE"& Chr(129) & document.all.TxtMCTYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CARRIER_ID"& Chr(129) & document.all.CARRIER_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "AGENT_ID"& Chr(129) & document.all.AGENT_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TPA_ID"& Chr(129) & document.all.TPA_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDITIONAL_DELIVERIES" & Chr(129) & document.all.ADDITIONAL_DELIVERIES.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DIVISION_CD"& Chr(129) & document.all.TxtDivisionCD.value & Chr(129) & "1" & Chr(128)


		if document.all.ChkSelfInsured.checked = True then
			sResult = sResult & "SELF_INSURED"& Chr(129) & "Y"  & Chr(129) & "1" & Chr(128)
		else
			sResult = sResult & "SELF_INSURED"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if

		sResult = sResult & "EFFECTIVE_DATE"& Chr(129) & "to_date('" & document.all.TxtEffective.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult = sResult & "EXPIRATION_DATE"& Chr(129) & "to_date('" & document.all.TxtExpiration.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult = sResult & "LOAD_DATE"& Chr(129) & "to_date('" & document.all.TxtLoad.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult = sResult & "CANCELLATION_DATE"& Chr(129) & "to_date('" & document.all.TxtCancellation.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult = sResult & "CHANGE_DATE"& Chr(129) & "to_date('" & document.all.TxtChange.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ORIGINAL_EFFECTIVE_DATE"& Chr(129) & "to_date('" & document.all.TxtOriginalEffective.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)

       	' Only If its an insert , we need the column ahs_policy_id
		If document.all.PID.value = "NEW" then
		   sResult1 = sResult1 & "AHS_POLICY_ID"& Chr(129)  & "1" & Chr(128)
		end if
		sResult1 = sResult1 & "Policy_ID"& Chr(129) & document.all.PID.value & Chr(129) & "1" & Chr(128)
		sResult1 = sResult1 & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult1 = sResult1 & "LOB_CD"& Chr(129) & document.all.TxtLOBCD.value & Chr(129) & "1" & Chr(128)
		sResult1 = sResult1 & "ACTIVE_START_DT"& Chr(129) & "to_date('" & document.all.TxtEffective.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult1 = sResult1 & "ACTIVE_END_DT"& Chr(129) & "to_date('" & document.all.TxtExpiration.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)

		'MMAI-0007 FOR POLICY_EXTENSION TABLE
		'Prashant Shekhar 05/22/2007

		If document.all.PID.value = "NEW" then
		   sResult2= sResult2 & "POLICY_EXTENSION_ID"& Chr(129)  & "1" & Chr(128)
		end if
		sResult2 = sResult2 & "Policy_ID"& Chr(129) & document.all.PID.value & Chr(129) & "1" & Chr(128)
		sResult2 = sResult2 & "NAME"& Chr(129) & "CLAIM:POLICY:CONTRACT_NUMBER" & Chr(129) & "1" & Chr(128)
		sResult2 = sResult2 & "VALUE"& Chr(129) & document.all.TxtContractNo.value & Chr(129) & "1" & Chr(128)


		document.all.TxtSaveData1.Value = sResult
		document.all.TxtSaveData2.Value = sResult1
		document.all.TxtSaveData3.Value = sResult2
		FrmDetails.action = "PolicySave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"
		FrmDetails.submit

		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If

	ExeSave = bRet
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"
	end if
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
<!--#include file="..\lib\Help.asp"-->
</script>
<script LANGUAGE="JavaScript" FOR="JurisBtnControl" EVENT="onscriptletevent (event, obj)">
	switch (event)
	{
		case "ATTACHBUTTONCLICK":
				ExeButtonsAttachJuris();
			break;
		default:
			break;
	}

</script>
<script LANGUAGE="JavaScript" FOR="CoverageBtnControl" EVENT="onscriptletevent (event, obj)">
  switch (event)
	{
		case "EDITBUTTONCLICK":
			EditCoverage();
			break;

		case "NEWBUTTONCLICK":
			NewCoverage();
			break;

		case "REMOVEBUTTONCLICK":
			RemoveCoverage();
			break;
		default:
			break;
	}

</script>
<!-- MMAI-0007
	 Prashant Shekhar 05/21/2007
	 The javascript onlyNumbers below doesnot allow the user to type in any other values other
	 than numbers. This is required for the COntract Number field which will take only valid
	 4 digit numbers. -->
<script ID = "ForNumbers" LANGUAGE="JavaScript">
function onlyNumbers()
 {
 //Remove numeric check as per KFAB-4409 -- 01 March 2010
 //if(event.keyCode < 48 || event.keyCode > 57)
 // event.returnValue = false;
 //else if(event.which < 48 || event.which > 57)
 //{
 // return false;
 // }
   Control_OnChange();
 }


</script>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Policy Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="PolicySave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData1">
<input TYPE="HIDDEN" NAME="TxtSaveData2">
<input TYPE="HIDDEN" NAME="TxtSaveData3">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="HIDDEN" NAME="TxtAHSID">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchPID" value="<%=Request.QueryString("SearchPID")%>">
<input type="hidden" name="SearchNumber" value="<%=Request.QueryString("SearchNumber")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchCarrier" value="<%=Request.QueryString("SearchCarrier")%>">
<input type="hidden" name="SearchTPADMIN" value="<%=Request.QueryString("SearchTPADMIN")%>">
<input type="hidden" name="SearchAgent" value="<%=Request.QueryString("SearchAgent")%>">
<input type="hidden" name="SearchLOBCD" value="<%=Request.QueryString("SearchLOBCD")%>">
<input type="hidden" name="SearchMCTYPE" value="<%=Request.QueryString("SearchMCTYPE")%>">
<input type="hidden" name="SearchSelfInsuredFlg" value="<%=Request.QueryString("SearchSelfInsuredFlg")%>">
<input type="hidden" name="SearchEffective" value="<%=Request.QueryString("SearchEffective")%>">
<input type="hidden" name="SearchOriginalEffective" value="<%=Request.QueryString("SearchOriginalEffective")%>">
<input type="hidden" name="SearchExpiration" value="<%=Request.QueryString("SearchExpiration")%>">
<input type="hidden" name="SearchCancellation" value="<%=Request.QueryString("SearchCancellation")%>">
<input type="hidden" name="SearchChange" value="<%=Request.QueryString("SearchChange")%>">
<input type="hidden" name="SearchLoad" value="<%=Request.QueryString("SearchLoad")%>">
<input type="hidden" name="SearchCompanyCode" value="<%=Request.QueryString("SearchCompanyCode")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PID" value="<%=Request.QueryString("PID")%>">
<input type="hidden" NAME="LOB" value="<%=Request.QueryString("LOB")%>">
<!-- MMAI -0007 -->
<input type="hidden" NAME="CLIENTAHSID" value="<%=Request.QueryString("AHSID")%>">
<%

IF PID = "NEW" then
Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
SQLST = "SELECT client_node_id clientid,parent_node_id from ACCOUNT_HIERARCHY_STEP where ACCNT_HRCY_STEP_ID= '"& trim(CLIENTAHSID) &"'"
Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			if RS("parent_node_id") = "1" then
				RSCLIENTNODEID = trim(CLIENTAHSID)
			else
				RSCLIENTNODEID = RS("clientid")
			end if
		end if
	rs.close
	set rs=nothing
end if

If PID <> "" Then
	If PID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING

		'**************************************************
		' DMS: 2/17/00 Changed the SQL to grab the columns
		'              LOB_CD and ACCNT_HRCY_STEP_ID from
		'              AHS_POLICY as the columns have been removed from
		'              the POLICY table.
		'**************************************************

		'********************MMAI- 0007******************************
		' Prashant Shekhar : 05/21/2007 Changed the SQL to grab the columns
		'              CLIENT_NODE_ID and PARENT_NODE_ID from
		'              ACCOUNT_HIERARCHY_STEP.
		'**************************************************

		SQLST = "SELECT " &_
		        " POLICY.*, " &_
				" AHS_POLICY.ACCNT_HRCY_STEP_ID, " &_
				" AHS_POLICY.LOB_CD, " &_
				" (SELECT  VALUE FROM POLICY_EXTENSION WHERE POLICY_EXTENSION.POLICY_ID = POLICY.Policy_ID " &_
				" AND NAME = 'CLAIM:POLICY:CONTRACT_NUMBER') AS CONTRACT_NO, " &_
				" to_char(LOAD_DATE, 'MM/DD/YYYY') F_LOAD_DATE, " &_
				" to_char(CANCELLATION_DATE, 'MM/DD/YYYY') F_CANCELLATION_DATE, " &_
				" to_char(EXPIRATION_DATE, 'MM/DD/YYYY') F_EXPIRATION_DATE, " &_
				" to_char(ORIGINAL_EFFECTIVE_DATE, 'MM/DD/YYYY') F_ORIGINAL_EFFECTIVE_DATE, " &_
				" to_char(EFFECTIVE_DATE, 'MM/DD/YYYY') F_EFFECTIVE_DATE, " &_
				" to_char(CHANGE_DATE, 'MM/DD/YYYY') F_CHANGE_DATE, " &_
				" AGENT.NAME AGENTNAME, CARRIER.NAME CARRIERNAME, " &_
				" THIRD_PARTY_ADMINISTRATOR.NAME TPANAME, " &_
				" ACCOUNT_HIERARCHY_STEP.NAME AHSID_TEXT, " &_
				" ACCOUNT_HIERARCHY_STEP.CLIENT_NODE_ID CLIENT_ID, " &_
				" ACCOUNT_HIERARCHY_STEP.PARENT_NODE_ID PARENT_ID " &_
				"FROM POLICY, " &_
				"     AHS_POLICY, " &_
				"     AGENT, " &_
				"     CARRIER, " &_
				"     THIRD_PARTY_ADMINISTRATOR, " &_
				"     ACCOUNT_HIERARCHY_STEP " &_
				"WHERE " &_
				" POLICY.Policy_ID                     = " & PID & " AND " &_
				" POLICY.AGENT_ID               = AGENT.AGENT_ID(+) AND " &_
				" POLICY.POLICY_ID              = AHS_POLICY.POLICY_ID AND " &_
				" AHS_POLICY.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				" POLICY.CARRIER_ID             = CARRIER.CARRIER_ID(+) AND " &_
				" POLICY.TPA_ID             = THIRD_PARTY_ADMINISTRATOR.TPA_ID(+) "

		'RESPONSE.WRITE(sqlst)


		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSCARRIER = RS("CARRIER_ID")
			RSTPAID = RS("TPA_ID")
			RSAGENT = RS("AGENT_ID")
			RSCARRIERNAME = ReplaceQuotesInText(RS("CARRIERNAME"))
			RSTPANAME = ReplaceQuotesInText(RS("TPANAME"))
			RSAGENTNAME = ReplaceQuotesInText(RS("AGENTNAME"))
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
			RSAHSID_TEXT = ReplaceQuotesInText(RS("AHSID_TEXT"))
			RSNUMBER = RS("POLICY_NUMBER")
			RSDESCRIPTION = ReplaceQuotesInText(RS("POLICY_DESC"))
			'MMAI-0023
			RSPOLICY_TYPE = RS("POLICY_TYPE")
			'End MMAI-0023
			RSLOBCD = RS("LOB_CD")
			RSCOMPANYCODE = RS("COMPANY_CODE")
			RSEFFECTIVE = RS("F_EFFECTIVE_DATE")
			RSORIGINALEFFECTIVE = RS("F_ORIGINAL_EFFECTIVE_DATE")
			RSDIVISION_CD = RS("DIVISION_CD")
			RSEXPIRATION = RS("F_EXPIRATION_DATE")
			RSCANCELLATION = RS("F_CANCELLATION_DATE")
			RSCHANGE = RS("F_CHANGE_DATE")
			RSLOAD = RS("F_LOAD_DATE")
			RSSELFINSUREDFLG = RS("SELF_INSURED")
			RSMCTYPE = RS("MANAGED_CARE_TYPE")
			RSADDITIONAL_DELIVERIES = RS("ADDITIONAL_DELIVERIES")
			RSCONTRACT_NUMBER = RS("CONTRACT_NO")
			RSCLIENTNODEID = RS("CLIENT_ID")
			RSPARENTNODEID = RS("PARENT_ID")
			if RSPARENTNODEID = "1" then
				RSCLIENTNODEID = RSAHSID
			else
				RSCLIENTNODEID = RS("CLIENT_ID")
			end if
		End If

		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing

	End If
%>

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

<table CLASS="LABEL" ID = "tblPolicy">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td>Policy ID:&nbsp;<span id="spanPID"><%=Request.QueryString("PID")%></span></td></tr>
<tr>
<td colspan="2">Number:<br><input ScrnInput="TRUE" MAXLENGTH="40" CLASS="LABEL" size="45" TYPE="TEXT" NAME="TxtNumber" VALUE="<%=RSNUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan="1">Description:<br><input ScrnInput="TRUE" MAXLENGTH="80" CLASS="LABEL" size="45" TYPE="TEXT" NAME="TxtDescription" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan="1">Policy Type:<br>
 <select ScrnBtn="TRUE" NAME="selPolicyType" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange">
 <option VALUE="C" <%if RSPOLICY_TYPE="C" OR RSPOLICY_TYPE="" then Response.write "selected"%>>CLAIM
 <option VALUE="R" <%if RSPOLICY_TYPE="R" then Response.write "selected"%>>RECORD ONLY
 <option VALUE="A" <%if RSPOLICY_TYPE="A" then Response.write "selected"%>>ALL POLICY TYPES
</select></td>
</tr>
<tr>
<td>Company Code:<br><input ScrnInput="TRUE" MAXLENGTH="6" CLASS="LABEL" size="8" TYPE="TEXT" NAME="TxtCompanyCode" VALUE="<%=RSCOMPANYCODE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>LOB:<br><select ScrnBtn="TRUE" NAME="TxtLOBCD" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOBCD,true)%></select></td>
<td>Managed Care Type:<br><select ScrnBtn="TRUE" NAME="TxtMCTYPE" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange">
 <option VALUE="NEITHER">NEITHER
 <option VALUE="CERTIFIED">CERTIFIED
 <option VALUE="NOTCERTIFIED">NOTCERTIFIED
</select></td>
<td><input ScrnBtn="TRUE" TYPE="CHECKBOX" NAME="ChkSelfInsured" ONCLICK="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" <% If CStr(RSSELFINSUREDFLG) = "Y" Then Response.Write("CHECKED")%>>Self Insured?</td>
</tr>

<tr>
<td>Effective Date:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtEffective" VALUE="<%=RSEFFECTIVE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Expiration Date:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtExpiration" VALUE="<%=RSEXPIRATION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Load Date:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtLoad" VALUE="<%=RSLOAD%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td>Cancellation Date:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtCancellation" VALUE="<%=RSCANCELLATION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Change Date:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtChange" VALUE="<%=RSCHANGE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Orig. Effective Date:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtOriginalEffective" VALUE="<%=RSORIGINALEFFECTIVE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr id = "trPolicy"><td>Division CD:<br><input ScrnInput="TRUE" MAXLENGTH="30" CLASS="LABEL" TYPE="TEXT" NAME="TxtDivisionCD" VALUE="<%=RSDIVISION_CD%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<!--MMAI-0007
 Prashant Shekhar 05/22/2007
 Contract is to be displayed for all clients.  -->
<td id = "tdContract">Contract No:<br><input ScrnInput="TRUE" MAXLENGTH="4" CLASS="LABEL" TYPE="TEXT" ID = "ContractNo" NAME="TxtContractNo" VALUE="<%=RSCONTRACT_NUMBER%>" ONKEYPRESS="javascript:onlyNumbers();" ONCHANGE="VBScript::Control_OnChange" ID="Text2"></td>
</tr>
</table>

<table class="LABEL">
<tr>
	<td><nobr>
	<img NAME="BtnAttachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach
	" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<img NAME="BtnDetachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	</td>
	<td width="325" nowrap><nobr>Account:&nbsp;<span ID="AHSID_TEXT" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(RSAHSID_TEXT)%>"><%=TruncateText(RSAHSID_TEXT,NameTextLen)%></span></td>
	<td>A.H.Step ID:&nbsp;<span ID="AHSID_ID" CLASS="LABEL"><%=RSAHSID%></span></td>
	</tr>
</table>

<table class="Label">
<td>
<img STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Carrier" ONCLICK="VBScript::AttachCarrier CARRIER_ID, CARRIER_NAME">
<img STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Carrier" OnClick="VBScript::Detach  CARRIER_ID, CARRIER_NAME">
</td>
<td width="325" nowrap>Carrier Name:&nbsp;<span ID="CARRIER_NAME" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(RSCARRIERNAME)%>"><%=TruncateText(RSCARRIERNAME,NameTextLen)%></span></td>
<td>Carrier ID:&nbsp;<span ID="CARRIER_ID" CLASS="LABEL"><%=RSCARRIER%></span></td>
</table>

<table class="Label">
<td>
<img NAME="BtnAttach" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Agent" ONCLICK="VBScript::AttachAgent AGENT_ID, AGENT_NAME">
<img NAME="BtnDetach" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Agent" OnClick="VBScript::Detach AGENT_ID, AGENT_NAME">
</td>
<td width="325" nowrap>Agent Name:&nbsp;<span ID="AGENT_NAME" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(RSAGENTNAME)%>"><%=TruncateText(RSAGENTNAME,NameTextLen)%></span></td>
<td>Agent ID:&nbsp;<span ID="AGENT_ID" CLASS="LABEL"><%=RSAGENT%></span></td>
</table>

<table class="Label" ID="Table1">
<td>
<img  STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach TPA" ONCLICK="VBScript::AttachTPA TPA_ID, TPA_NAME">
<img  STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach TPA" OnClick="VBScript::Detach TPA_ID, TPA_NAME">
</td>
<td width="325" nowrap>T P A  Name:&nbsp;<span ID="TPA_NAME" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(RSTPANAME)%>"><%=TruncateText(RSTPANAME,NameTextLen)%></span></td>
<td>T P A  ID:&nbsp;<span ID="TPA_ID" CLASS="LABEL"><%=RSTPAID%></span></td>
</table>

<table WIDTH="100%">
<tr>
<td CLASS="LABEL">Additional Deliveries:<br>
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="ADDITIONAL_DELIVERIES" CLASS="LABEL" SIZE="76" MAXLENGTH="255" VALUE="<%= RSADDITIONAL_DELIVERIES %>">
<img SRC="../Images/PropertiesIcon.gif" ID="BtnEDITDELIVERIES" STYLE="CURSOR:HAND" ALT="Edit Additional Deliveries" OnClick="EditDeliveries(ADDITIONAL_DELIVERIES)" WIDTH="16" HEIGHT="14"></td>
</tr>
</table>

<table>
<td>

<table CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Jurisdiction States</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="50" HEIGHT="8"></td></tr>
</table></td></tr>

<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset ID="fldSet1" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'0';width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&amp;ATTACHCAPTION=Edit&amp;HIDEREMOVE=TRUE&amp;HIDENEW=TRUE&amp;HIDEEDIT=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="JurisBtnControl" type="text/x-scriptlet" VIEWASTEXT></object>
<iframe FRAMEBORDER="0" height="0" width="100%" align="left" name="JurisFrame" src="PolicyDetailsJuris.asp?<%=Request.QueryString%>" scrolling="auto"></iframe>
</fieldset>

<table CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="175" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>

</table>
<!--<fieldset ID="fldSet2" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'0';width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&amp;HIDEATTACH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="CoverageBtnControl" type="text/x-scriptlet"></object>
<iframe FRAMEBORDER="0" name="CoverageFrame" width="100%" height="0" src="PolicyDetailsCoverage.asp?<%=Request.QueryString%>" scrolling="auto"></iframe>

</fieldset>
</td>-->
</table>

<%	If Not IsNull(RSMCTYPE) Then
		If  CStr(RSMCTYPE) <> "" Then %>
<script LANGUAGE="VBScript">
	SelectOption document.all.TxtMCTYPE,"<%=CStr(RSMCTYPE)%>"
</script>
<%		End If
	End If%>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Policy selected.
</div>

<% End If %>

</form>
</body>
</html>