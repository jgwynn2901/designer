<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<% Response.Expires = 0 
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CNodeSearchObj()
{
	this.AHSID = "";
	this.Selected = false;
}

function COutputDefinitionSearchObj()
{
	this.ODID = "";
	this.ODIDName = "";
	this.Saved = false;	
	this.Selected = false;	
}

function CRoutingRelatedSearchObj()
{
	this.RPID = "";
	this.RPDesc = "";
	this.RPSelected = false;
}

var RoutingPlanObj = new CRoutingRelatedSearchObj();
var DefinitionObj = new COutputDefinitionSearchObj();
var NodeSearchObj = new CNodeSearchObj();
var g_StatusInfoAvailable = false;
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->
dim nOffset

Sub BtnShowSettings_onclick
If document.all.DIVCUSTOM.Style.display = "none" Then
	document.all.DIVCUSTOM.Style.display = "block"
Else
	document.all.DIVCUSTOM.Style.display = "none"
End If
End Sub

Sub AllRoutingRelatedItems_OnClick
	IF document.all.AllRoutingRelatedItems.Checked = False THEN
		document.all.Selected_RoutingRelatedItems.style.display = "block"
	ELSE
		document.all.Selected_RoutingRelatedItems.style.display = "none"
	END IF
End Sub

Sub CLIENTROUTING_OnClick
	IF document.all.CLIENTROUTING.Checked = True THEN
		document.all.AllRoutingRelatedItems.disabled = False
		If document.all.AllRoutingRelatedItems.Checked = False Then
			document.all.Selected_RoutingRelatedItems.style.display = "block"
		End If
	ELSE
		If document.all.AllRoutingRelatedItems.Checked = False Then
			document.all.Selected_RoutingRelatedItems.Style.Display = "none"
		End If
		document.all.AllRoutingRelatedItems.disabled = True
	END IF
End Sub

Sub ALLDEFINITIONS_OnClick
If document.all.ALLDEFINITIONS.checked = false Then
	document.all.SELECTED_OUTPUTDEFS.style.display="block"
Else
	document.all.SELECTED_OUTPUTDEFS.style.display="none"
End If
End Sub

Sub OUTPUTDEFINITION_OnClick
	If document.all.OUTPUTDEFINITION.checked = true then
		document.all.ALLDEFINITIONS.disabled = false
		If document.all.ALLDEFINITIONS.checked = false Then
			document.all.SELECTED_OUTPUTDEFS.style.display="block"
		End If
	Else
		If document.all.ALLDEFINITIONS.checked = false Then
			document.all.SELECTED_OUTPUTDEFS.style.display="none"
		End If
		document.all.ALLDEFINITIONS.disabled = true
	End If
End Sub

Function Validate
    Dim cErrMsg, cTime, lToday

    cErrMsg = ""
    lToday = True
    If document.All.AHSID.Value = "" Then
        cErrMsg = "A.H.S. ID is a required field." & vbCrLf
    End If
    If document.All.LOB_CD.Value = "" Then
        cErrMsg = cErrMsg & "LOB is a required field" & vbCrLf
    End If
	If document.All.REFERENCE_ID.Value = "" Then
        cErrMsg = cErrMsg & "Reference ID is a required field." & vbCrLf
    End If
    If document.All.START_DATE.Value = "" Then
        cErrMsg = cErrMsg & "Please enter a start date." & vbCrLf
    ElseIf Not CheckDate(document.All.START_DATE.Value) Then
        cErrMsg = cErrMsg & "Invalid start date, use mm/dd/yy format." & vbCrLf
    ElseIf CDate(document.All.START_DATE.Value) < Date Then
        cErrMsg = cErrMsg & "Start date must be today or a future date." & vbCrLf
    ElseIf (document.All.START_DATE.Value <> FormatDateTime(Date, 2)) And (document.All.START_TIME.Value = "") Then
        cErrMsg = cErrMsg & "If you choose a date other than today, you MUST enter a valid time." & vbCrLf
    End If

    If cErrMsg = "" Then
        If CDate(document.All.START_DATE.Value) > Date Then
            lToday = False
        End If
    End If
    If document.All.START_TIME.Value <> "" Then
        If Not checkTime(document.All.START_TIME.Value) Then
            cErrMsg = cErrMsg & "Invalid Start time, use hh:mm format." & vbCrLf
        End If
    End If
    If document.All.RULES.Checked = False And document.All.RULES_LOOKUPS.Checked = False And _
            document.All.CLAIMNUMBER.Checked = False And document.All.Attributes.Checked = False And _
            document.All.ASSIGNMENT.Checked = False And document.All.ACCOUNTCALLFLOW.Checked = False And _
            document.All.ROUTINGADDRESS.Checked = False And document.All.ESCALATION_RULES.Checked = False And _
            document.All.Information.Checked = False And _
            document.All.COMMON.Checked = False And document.All.CLIENTROUTING.Checked = False And _
            document.All.EDIROUTING.Checked = False And document.All.OUTPUTDEFINITION.Checked = False And _
            document.All.OutputOverFlow.Checked = False And document.All.VendorDefs.Checked = False And _
            document.All.FraudDefs.Checked = False And document.All.Subrogation.Checked = False Then
        cErrMsg = cErrMsg & "Please Select At Least One Option to Migrate." & vbCrLf
    End If

    cTime = FormatDateTime(CDate(FormatDateTime(Time(), vbShortTime)) - nOffset, vbShortTime)
    document.All.CurrTime.innerHTML = cTime
    If cErrMsg = "" Then
        If document.All.START_TIME.Value <> "" Then
            If CDate(document.All.START_TIME.Value) < CDate(cTime) And lToday Then
                cErrMsg = "If you want the Migration job to start right away, leave the time field empty."
                document.All.START_TIME.Value = ""
            End If
        End If
    End If
    If cErrMsg = "" Then
        If document.All.START_TIME.Value <> "" Then
            If CDate(document.All.START_TIME.Value) < DateAdd("n", 5, CDate(cTime)) And lToday Then
                cErrMsg = "The Start time is too close to the current time." & vbCrLf & "Leave it blank to start immediately or schedule at least 5 minutes in advance."
            End If
        End If
    End If
    If cErrMsg = "" Then
        Validate = True
    Else
        MsgBox cErrMsg, 0, "FNSDesigner"
        Validate = False
    End If
End Function

Function Swap (InData)
If InData = "" Then
	Swap = "null"
Else
	Swap = InData
End If
End Function

Sub BtnMigrate_onclick
dim cMsg, migrationStatus
RPList = ""
ODlist = ""
sResult = ""
migrationStatus="true"


If Validate() Then
    If( ( "<%=Session("isAsp")%>" = false AND "<%=Session("ENVIRONMENT_ABBREVIATION")%>" = "QA") OR "<%=Session("ENVIRONMENT_ABBREVIATION")%>" = "PP") Then
         migrationStatus="false"
         migrationStatus=window.showModalDialog("MigrationConfirmationDialog.asp","","dialogwidth:510px;dialogheight:130px")
    End If
    If migrationStatus="true" Then
        cMsg = "Are you sure you wish to migrate this data? This cannot be undone!" & vbCrLf
        If document.All.chkLevelOne.Checked Then
            cMsg = cMsg & vbCrLf & "[ Level ONE Migration ]" & vbCrLf
        End If
        If document.All.RULES.Checked = True Then
            cMsg = cMsg & "Branch Rules" & vbCrLf
        End If
        If document.All.RULES_LOOKUPS.Checked = True Then
            cMsg = cMsg & "Rules and Lookups" & vbCrLf
        End If
        If document.All.CLAIMNUMBER.Checked = True Then
            cMsg = cMsg & "Claim Number" & vbCrLf
        End If
        If document.All.Attributes.Checked = True Then
            cMsg = cMsg & "Attributes" & vbCrLf
        End If
        If document.All.ASSIGNMENT.Checked = True Then
            cMsg = cMsg & "Assignment" & vbCrLf
        End If
        If document.All.ACCOUNTCALLFLOW.Checked = True Then
            cMsg = cMsg & "Account Call Flow" & vbCrLf
        End If
        If document.All.ROUTINGADDRESS.Checked = True Then
            cMsg = cMsg & "Routing Address" & vbCrLf
        End If
        If document.All.ESCALATION_RULES.Checked = True Then
            cMsg = cMsg & "Escalation Rules" & vbCrLf
        End If
        If document.All.Information.Checked = True Then
            cMsg = cMsg & "Information" & vbCrLf
        End If
        If document.All.COMMON.Checked = True Then
            cMsg = cMsg & "Common " & vbCrLf
        End If
        If document.All.OutputOverFlow.Checked = True Then
            cMsg = cMsg & "Output OverFlow " & vbCrLf
        End If
        If document.All.CLIENTROUTING.Checked = True And document.All.ALLROUTINGRELATEDITEMS.Checked = True Then
            cMsg = cMsg & "All Routing Plan Related" & vbCrLf
        End If
        If document.All.CLIENTROUTING.Checked = True And document.All.ALLROUTINGRELATEDITEMS.Checked = False Then
            cMsg = cMsg & "Selected Routing Related Items" & vbCrLf
        End If
        If document.All.EDIROUTING.Checked = True Then
            cMsg = cMsg & "EDI UOF Routing Related" & vbCrLf
        End If
        If document.All.ALLDEFINITIONS.Checked = True And document.All.OUTPUTDEFINITION.Checked = True Then
            cMsg = cMsg & "All Output Definitions" & vbCrLf
        End If
        If document.All.ALLDEFINITIONS.Checked = False And document.All.OUTPUTDEFINITION.Checked = True Then
            cMsg = cMsg & "Selected Output Definitions" & vbCrLf
        End If
        If document.All.VendorDefs.Checked Then
            cMsg = cMsg & "Vendor Eligibility Definitions" & vbCrLf
        End If
        If document.All.FraudDefs.Checked Then
            cMsg = cMsg & "Fraud Detection Definitions" & vbCrLf
        End If
        If document.All.Subrogation.Checked Then
            cMsg = cMsg & "Subrogation Definitions" & vbCrLf
        End If
        lret = MsgBox(cMsg, 1, "FNSDesigner")

        If lret = "1" Then
            If document.All.Start_Time.Value = "" Then
                document.All.Start_Time.Value = DatePart("h", Time) & ":" & DatePart("n", Time)
            End If
            sResult = sResult & "JOB_ID" & Chr(129) & "" & Chr(129) & "0" & Chr(128)
            sResult = sResult & "SCHEDULED_START" & Chr(129) & "TO_DATE('" & document.All.Start_Date.Value & "," & document.All.Start_Time.Value & ":" & Second(Now) & "', 'MM/DD/YY,HH24:MI:SS')" & Chr(129) & "0" & Chr(128)
            sResult = sResult & "ACCNT_HRCY_STEP_ID" & Chr(129) & document.All.AHSID.Value & Chr(129) & "1" & Chr(128)
            sResult = sResult & "LOB_CD" & Chr(129) & document.All.LOB_CD.Value & Chr(129) & "1" & Chr(128)
			sResult = sResult & "REFERENCE_ID" & Chr(129) & document.All.REFERENCE_ID.Value & Chr(129) & "1" & Chr(128)
            sResult = sResult & "STATUS_CD" & Chr(129) & "1" & Chr(129) & "1" & Chr(128)
            sResult = sResult & "STATUS_MSG" & Chr(129) & "null" & Chr(129) & "0" & Chr(128)
            sResult = sResult & "USER_ID" & Chr(129) & "<%= Session("SecurityObj").m_UserID %>" & Chr(129) & "0" & Chr(128)
            If document.All.chkLevelOne.Checked Then
                sResult = sResult & "LEVEL_ONE" & Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
            Else
                sResult = sResult & "LEVEL_ONE" & Chr(129) & "N" & Chr(129) & "1" & Chr(128)
            End If
            If document.All.VendorDefs.Checked Then
                sResult = sResult & "VENDOR_ELIGIBILITY" & Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
            Else
                sResult = sResult & "VENDOR_ELIGIBILITY" & Chr(129) & "N" & Chr(129) & "1" & Chr(128)
            End If
            If document.All.FraudDefs.Checked Then
                sResult = sResult & "FRAUD_DETECTION" & Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
            Else
                sResult = sResult & "FRAUD_DETECTION" & Chr(129) & "N" & Chr(129) & "1" & Chr(128)
            End If
            If document.All.ALLROUTINGRELATEDITEMS.Checked = True And document.All.CLIENTROUTING.Checked = True Then
                sResult = sResult & "MOVE_ALL_ROUTING_PLANS" & Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
            Else
                sResult = sResult & "MOVE_ALL_ROUTING_PLANS" & Chr(129) & "N" & Chr(129) & "1" & Chr(128)
            End If
            If document.All.ALLDEFINITIONS.Checked = True And document.All.OUTPUTDEFINITION.Checked = True Then
                sResult = sResult & "MOVE_ALL_OUTPUT_DEFS" & Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
            Else
                sResult = sResult & "MOVE_ALL_OUTPUT_DEFS" & Chr(129) & "N" & Chr(129) & "1" & Chr(128)
            End If
            sResult = sResult & "START_TIME" & Chr(129) & "null" & Chr(129) & "0" & Chr(128)
            sResult = sResult & "END_TIME" & Chr(129) & "null" & Chr(129) & "0" & Chr(128)

            For R = 0 To document.All.RoutingRelatedItems_List.Options.length - 1 Step 1
                If RPList <> "" Then
                    RPList = RPList & ","
                End If
                RPList = RPList & document.All.RoutingRelatedItems_List.Options(R).Value
            Next

            For i = 0 To document.All.OUTPUTDEFIDS_LIST.Options.length - 1 Step 1
                If ODlist <> "" Then
                    ODlist = ODlist & ","
                End If
                ODlist = ODlist & document.All.OUTPUTDEFIDS_LIST.Options(i).Value
            Next

            If document.All.CLIENTROUTING.Checked = True Then
                If document.All.ALLROUTINGRELATEDITEMS.Checked = True Then
                    document.All.HiddenALlRoutingRelatedItems.Value = ""
                Else
                    document.All.HiddenALlRoutingRelatedItems.Value = RPList
                End If
            Else
                document.All.HiddenALlRoutingRelatedItems.Value = ""
            End If

            If document.All.OUTPUTDEFINITION.Checked = True Then
                If document.All.ALLDEFINITIONS.Checked = True Then
                    document.All.HiddenODlist.Value = ""
                Else
                    document.All.HiddenODlist.Value = ODlist
                End If
            Else
                document.All.HiddenODlist.Value = ""
            End If

            document.All.TxtSaveData.Value = sResult
            FrmData.submit()
        End If
    End If
End If
End Sub

Function AttachNode(ID)
    AHSID = ID.Value
    MODE = document.body.getAttribute("ScreenMode")

    NodeSearchObj.AHSID = AHSID
    NodeSearchObj.Selected = False

    If AHSID = "" Then AHSID = "NEW"

    If AHSID = "NEW" And MODE = "RO" Then
        MsgBox "No AHS currently attached.", 0, "FNSNetDesigner"
        Exit Function
    End If

    strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_DATA_MIGRATION&SELECTONLY=TRUE&AHSID=" & AHSID
    If MODE = "RO" Then strURL = strURL & "&DETAILONLY=TRUE"

    showModalDialog strURL, NodeSearchObj, "dialogWidth=650px; dialogHeight=700px; center=yes"
    If NodeSearchObj.AHSID <> ID.Value Then
        document.body.setAttribute "ScreenDirty", "YES"
        ID.Value = NodeSearchObj.AHSID
    End If
End Function

FUNCTION f_RemoveRoutingPlan()
	RP_Index = ""
	For i = 0 to document.all.RoutingRelatedItems_List.options.length - 1 step 1
		IF document.all.RoutingRelatedItems_List.Options(i).Selected = True Then
			RP_Index = i
		END IF
	Next
	If RP_Index <> "" Then
		document.all.RoutingRelatedItems_List.Remove(RP_Index)
	End If
End FUNCTION

FUNCTION f_AttachRoutingPlan()
	l_Ret = Window.showModalDialog("../RoutingPlan/RoutingPlanSearchModal.asp?CONTAINERTYPE=FRAMEWORK&SELECTONLY=TRUE", RoutingPlanObj,"dialogWidth:650px;dialogHeight:450px;center")
	IF Not RoutingPlanObj.RPSelected THEN Exit Function
	IF RoutingPlanObj.RPID <> "" THEN
		If Instr(1, RoutingPlanObj.RPID, ",") Then
			Dim a_RPIDArray, a_RPDescArray
			a_RPIDArray = Split(RoutingPlanObj.RPID, ",")
			a_RPDescArray = Split(RoutingPlanObj.RPDesc, ",")
			sub_Add_MultiSelected "RP", a_RPIDArray, a_RPDescArray
		Else
			if document.all.RoutingRelatedItems_List.length > 0 then
				For j = 0 to document.all.RoutingRelatedItems_LIST.Options.length -1
					if RoutingPlanObj.RPID = document.all.RoutingRelatedItems_LIST.Options(j).Value Then Exit Function
				Next
			end if
			Set o_RPOption = document.createElement("option")
			o_RPOption.value = RoutingPlanObj.RPID
			o_RPOption.Text = "(" & RoutingPlanObj.RPID & ")" & RoutingPlanObj.RPDesc
			document.all.RoutingRelatedItems_List.Add(o_RPOption)
		End If
	END IF
END FUNCTION

Function RemoveODID()
	index = ""
	for i = 0 to document.all.OUTPUTDEFIDS_LIST.options.length-1 step 1
		if document.all.OUTPUTDEFIDS_LIST.options(i).selected = True Then
			index = i
		End If
	next
	If index <> "" Then
		document.all.OUTPUTDEFIDS_LIST.Remove(index)
	End If
End Function

Function AttachODID()
    lret = window.showModalDialog("../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=MODAL&SELECTONLY=TRUE", DefinitionObj, "dialogWidth:450px;dialogHeight:450px;center")
    If Not DefinitionObj.Selected Then
        Exit Function
    End If
    If DefinitionObj.ODID <> "" Then
        If InStr(1, DefinitionObj.ODID, "||") Then
            Dim a_IDArray, a_IDNameArray
            a_IDArray = Split(DefinitionObj.ODID, "||")
            a_IDNameArray = Split(DefinitionObj.ODIDName, "||")
            sub_Add_MultiSelected "OD", a_IDArray, a_IDNameArray
        Else
            If document.All.OUTPUTDEFIDS_LIST.length > 0 Then
                For j = 0 To document.All.OUTPUTDEFIDS_LIST.Options.length - 1
                    If DefinitionObj.ODID = document.All.OUTPUTDEFIDS_LIST.Options(j).Value Then Exit Function
                Next
            End If
            Set objOption = document.createElement("option")
            objOption.Value = DefinitionObj.ODID
            objOption.Text = "(" & DefinitionObj.ODID & ") " & DefinitionObj.ODIDName
            document.All.OUTPUTDEFIDS_LIST.Add (objOption)
        End If
    End If
End Function

Sub sub_Add_MultiSelected(s_ControlData, s_IDs, s_IDNames)
    If IsArray(s_IDs) Then
        Select Case s_ControlData
            Case "OD"
                For i = LBound(s_IDs) To UBound(s_IDNames)
                    b_PreviouslySelected = False
                    If document.All.OUTPUTDEFIDS_LIST.length > 0 Then       ' ListBox is not empty
                        For j = 0 To document.All.OUTPUTDEFIDS_LIST.Options.length - 1
                            If s_IDs(i) = Mid(document.All.OUTPUTDEFIDS_LIST.Options(j).Text, 2, InStr(1, document.All.OUTPUTDEFIDS_LIST.Options(j).Text, ")") - 2) Then
                                b_PreviouslySelected = True
                                Exit For
                            End If
                        Next
                    End If
                    If Not b_PreviouslySelected Then
                        Set objOption = document.createElement("option")
                        objOption.Value = s_IDs(i) 'DefinitionObj.ODID
                        objOption.Text = "(" & s_IDs(i) & ") " & s_IDNames(i) '"(" & DefinitionObj.ODID & ") " & DefinitionObj.ODIDName
                        document.All.OUTPUTDEFIDS_LIST.Add (objOption)
                    End If
                Next
            Case "RP"
                For i = LBound(s_IDs) To UBound(s_IDNames)
                    b_PreviouslySelected = False
                    If document.All.RoutingRelatedItems_LIST.length > 0 Then        ' ListBox is not empty
                        For j = 0 To document.All.RoutingRelatedItems_LIST.Options.length - 1
                            If s_IDs(i) = Mid(document.All.RoutingRelatedItems_LIST.Options(j).Text, 2, InStr(1, document.All.RoutingRelatedItems_LIST.Options(j).Text, ")") - 2) Then
                                b_PreviouslySelected = True
                                Exit For
                            End If
                        Next
                    End If
                    If Not b_PreviouslySelected Then
                        Set objOption = document.createElement("option")
                        objOption.Value = s_IDs(i) 'DefinitionObj.ODID
                        objOption.Text = "(" & s_IDs(i) & ") " & s_IDNames(i) '"(" & DefinitionObj.ODID & ") " & DefinitionObj.ODIDName
                        document.All.RoutingRelatedItems_LIST.Add (objOption)
                    End If
                Next
        End Select
    End If
End Sub

Sub RemoveOptions( objSelect )
	Dim intOptIndex
	If objSelect.options.length > 0 Then
		while (objSelect.options.length > 0)
			objSelect.Remove( intOptIndex )
		wend
	End if
End Sub

Sub Check_OnClick
If document.all.Check.innertext = "UnCheck All" Then
	document.all.Check.innertext = "Check All"
	UnCheckall()
Else
	document.all.Check.innertext = "UnCheck All"
	Checkall()
End If
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub window_onload
dim cServerTime, cClientTime

cServerTime = "<%=FormatDateTime(Time(),vbShortTime)%>"
cClientTime = FormatDateTime(Time(),vbShortTime)
nOffset = CDate(cClientTime) - CDate(cServerTime)
document.all.CurrTime.innerHTML = cServerTime
 
	CLIENTROUTING_OnClick()
	OUTPUTDEFINITION_OnClick()
	document.all.SpanStatus.innerHTML = "Ready"
	document.all.BtnShowSettings.Click()
	document.all.ROW1.Style.Cursor="HAND"
	document.all.ROW2.Style.Cursor="HAND"
	document.all.ROW3.Style.Cursor="HAND"
	document.all.ROW4.Style.Cursor="HAND"
	document.all.ROW5.Style.Cursor="HAND"
	document.all.ROW6.Style.Cursor="HAND"
	document.all.ROW7.Style.Cursor="HAND"
	document.all.ROW8.Style.Cursor="HAND"
	document.all.ROW10.Style.Cursor="HAND"
	document.all.ROW11.Style.Cursor="HAND"
	document.all.ROW12.Style.Cursor="HAND"
	document.all.ROW14.Style.Cursor="HAND"
	document.all.ROW15.Style.Cursor="HAND"
	document.all.ROW16.Style.Cursor="HAND"
	
	document.all.START_DATE.value = FormatDateTime(Date(),2)
End Sub

Sub CheckAll()
	document.all.RULES.checked = True 
	document.all.RULES_LOOKUPS.checked = True
	document.all.CLAIMNUMBER.checked = True
	document.all.ATTRIBUTES.checked = True
	document.all.ASSIGNMENT.checked = True
	document.all.ACCOUNTCALLFLOW.checked = True
	document.all.ROUTINGADDRESS.checked = True 
	document.all.ESCALATION_RULES.checked = True 
	document.all.INFORMATION.checked = True 
	'document.all.COMMON.checked = True -- DISABLED
	document.all.CLIENTROUTING.checked = True
	document.all.EDIROUTING.checked = True	
	document.all.OUTPUTDEFINITION.checked = True
	document.all.OUTPUTOVERFLOW.Checked =  True
	document.all.VendorDefs.Checked =  True
	document.all.FraudDefs.Checked =  True
	document.all.Subrogation.Checked =  True
	CLIENTROUTING_OnClick()
	OUTPUTDEFINITION_OnClick()
End Sub

Sub UnCheckAll()
	document.all.RULES.checked = False 
	document.all.RULES_LOOKUPS.checked = False
	document.all.CLAIMNUMBER.checked = False
	document.all.ATTRIBUTES.checked = False
	document.all.ASSIGNMENT.checked = False
	document.all.ACCOUNTCALLFLOW.checked = False
	document.all.ROUTINGADDRESS.checked = False 
	document.all.ESCALATION_RULES.checked = False 
	document.all.INFORMATION.checked = False 
	'document.all.COMMON.checked = False  -- DISABLED
	document.all.CLIENTROUTING.checked = False
	document.all.EDIROUTING.checked = False  
	document.all.OUTPUTDEFINITION.checked = False
	document.all.OUTPUTOVERFLOW.Checked = False
	document.all.VendorDefs.Checked =  False
	document.all.FraudDefs.Checked =  false
	document.all.Subrogation.Checked =  false
	CLIENTROUTING_OnClick()
	OUTPUTDEFINITION_OnClick()
End Sub

Function CheckDate(InDate)
dim sTemp, sMonth, sDay, sYear

	sTemp = split(inDate,"/")

	if(ubound(sTemp) < 2) then
		CheckDate = false
		exit function
	end if
	
	sMonth = sTemp(0)
	sDay = sTemp(1)
	sYear = sTemp(2)
	
	if (not IsNumeric(sMonth)) OR (not isnumeric(sDay)) or (not isnumeric(sYear)) then
		checkdate=false
		exit function
	end if
	
	if( sMonth < 1 or sMonth > 12) then
		checkdate = false
		exit function
	end if
	
	if(sDay < 1 or sDay > 31) then
		checkdate = false
		exit function
	end if
	
	CheckDate = True
End Function

Function checkTime(cTime)
dim cHours, cMin, nPos, nLen, x, cChar, lIsMin, lError
	
cTime = Trim( cTime )
nPos = InStr( 1, cTime, ":", vbTextCompare )
cHours = ""
cMin = ""
lIsMin = false
lError = false
if nPos <> 0 then
	nLen = Len( cTime )
	for x = 1 to nLen
		cChar = mid( cTime, x, 1 )
		if cChar = ":" then
			lIsMin = true
		else
			if isNumeric( cChar ) then
				if lIsMin then
					cMin = cMin & cChar
				else
					cHours = cHours & cChar
				end if
			else
				lError = true
				exit for
			end if
		end if
	next
else
	lError = true
end if
if not lError then
	if isNumeric( cHours ) and isNumeric( cMin ) then
		if not (cint(cHours) >= 0 and cint(cHours) <= 24 and cint(cMin) >= 0 and cint(cMin) <= 59) then
			lError = true
		end if
	else
		lError = true
	end if
end if
checkTime = not lError
End Function

-->
</script>
<script Language="JavaScript" For="BtnControl_RoutingPlanRelatedItems" Event="onscriptletevent (event, obj)">
	switch (event){
		case "REMOVEBUTTONCLICK":
			f_RemoveRoutingPlan()
			break;
		case "ATTACHBUTTONCLICK":
			f_AttachRoutingPlan()
			break;
		default:
			break;
	}
</script>
<script LANGUAGE="JavaScript" FOR="BtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
	case "REMOVEBUTTONCLICK":
		RemoveODID()
		break;
	case "ATTACHBUTTONCLICK":
		AttachODID()
		break;
	default:
		break;
}
 
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<form NAME="FrmData" method="post" ACTION="MigrationJobSave.asp">
	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
		<tr>
			<td colspan="2" HEIGHT="4"></td>
		</tr>
		<tr>
			<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Account Associated Control Data&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
			<td HEIGHT="5" ALIGN="LEFT">
				<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
					<tr>
						<td WIDTH="3" HEIGHT="4"></td>
						<td WIDTH="300" HEIGHT="4"></td>
					</tr>
					<tr>
						<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
						<td WIDTH="300" HEIGHT="8"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td>
		</tr>
		<tr>
			<td colspan="2" HEIGHT="1"></td>
		</tr>
	</table>
	<table class="Label" CELLSPACING="0" CELLPADDING="0">
		<tr>
			<td ALIGN="LEFT" VALIGN="BOTTOM"><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report"></td>
			<td ALIGN="LEFT" VALIGN="BOTTOM">:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span></td>
		</tr>
	</table>
	<input TYPE="HIDDEN" NAME="TxtAction" VALUE="INSERT">
	<input TYPE="HIDDEN" NAME="TxtSaveData">
	<input TYPE="HIDDEN" NAME="HiddenALlRoutingRelatedItems">
	<input TYPE="HIDDEN" NAME="HiddenODlist">
	<table WIDTH="100%">
		<tr>
			<td CLASS="LABEL" WIDTH="20"><nobr>LOB:<br>
				<select NAME="LOB_CD" CLASS="LABEL">
					<option VALUE>
					<option VALUE="ALL">All Lines of Business
					<%
					Set Conn = Server.CreateObject("ADODB.Connection")
					ConnectionString = CONNECT_STRING
					Conn.Open ConnectionString
					SQLST = ""
					SQLST = SQLST & "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
					Set RS = Conn.Execute(SQLST)
					Do While Not RS.EOF
					%>
					<option VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_NAME") %>
					<%
					RS.MoveNext
					Loop
					RS.CLose
					%>
				</select>
			</td>
			<td CLASS="LABEL" WIDTH="20"><nobr>A.H.S. ID:<br>
				<img NAME="BtnAttach" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach AHSID" OnClick="AttachNode(AHSID)">
				<input TYPE="TEXT" NAME="AHSID" CLASS="LABEL" STYLE="BACKGROUND-COLOR:SILVER" READONLY SIZE="10" MAXLENGTH="10">
			</td>
			<td CLASS="LABEL" TITLE="mm/dd/yy">Date:<br><input TYPE="TEXT" NAME="START_DATE" CLASS="LABEL" SIZE="12" MAXLENGTH="10"></td>
			<td CLASS="LABEL" TITLE="Enter the time in the 24 hour clock format ie. 17:30">Time<font SIZE="-5">(24 Hour Clock)</font>:<br><input TYPE="TEXT" NAME="START_TIME" CLASS="LABEL" SIZE="10" MAXLENGTH="5"></td>
			<td	VALIGN="BOTTOM" ALIGN="RIGHT"><button NAME="BtnMigrate" STYLE CLASS="STDBUTTON">Migrate</button>&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td> </td>
			<td> </td>
			<td> </td>
			<td CLASS="LABEL"><em><big>Current Time: <span ID="CurrTime"></span></big></em></td>
		</tr>
		<tr>
			<td><button NAME="BtnShowSettings" CLASS="STDBUTTON">Settings »</button></td>
		</tr>
	</table>
	<div ID="DIVCUSTOM" STYLE="DISPLAY:none">
	<table WIDTH="100%" CELLSPACING="2" CELLPADDING="1">
		<tr>
			<td CLASS="LABEL" BGCOLOR="#3399cc">Branch</td>
			<td CLASS="LABEL" BGCOLOR="#3399cc">Call flow</td>
		</tr>
		<tr>
			<td CLASS="LABEL" ID="ROW1" TITLE="Branch Rules, Claim Assignment Rules, Managed Care Rules"><input TYPE="CHECKBOX" NAME="RULES" CHECKED>Rules</td>
			<td CLASS="LABEL" ID="ROW2" TITLE="Valid Value, Rules, Location Code, LU Type, LU Code"><input TYPE="CHECKBOX" NAME="RULES_LOOKUPS" CHECKED>Rules/Look ups</td>
		</tr>
		<tr>
			<td CLASS="LABEL" ID="ROW3" TITLE="Claim Number Assignment Rules"><input TYPE="CHECKBOX" NAME="CLAIMNUMBER" CHECKED>Claim Number</td>
			<td CLASS="LABEL" ID="ROW4" TITLE="Attributes, Frame, Atrr Instance, Attrinbute Override"><input TYPE="CHECKBOX" NAME="ATTRIBUTES" CHECKED>Attributes</td>
		</tr>
		<tr>
			<td CLASS="LABEL" ID="ROW5" TITLE="Branch Assignment Type, Branch Assignment Rules, Managed Care Assignment Type, Managed Care Assignment Rules"><input TYPE="CHECKBOX" NAME="ASSIGNMENT" CHECKED>Assignment</td>
			<td CLASS="LABEL" ID="ROW6" TITLE="Call Flow, Frame Order, Address Book Entry, Inbound Call, Account Call Flow"><input TYPE="CHECKBOX" NAME="ACCOUNTCALLFLOW" CHECKED>Account Call Flow</td>
		</tr>
		<tr>
			<td CLASS="LABEL" ID="ROW7" TITLE="Routing Address, Routing Address Rule"><input TYPE="CHECKBOX" NAME="ROUTINGADDRESS" CHECKED>Routing Address</td>
		</tr>
		<tr>
			<td CLASS="LABEL" BGCOLOR="#3399cc">Escalation</td>
			<td CLASS="LABEL" BGCOLOR="#3399cc">Billing</font></td>
		</tr>
		<tr>
			<td CLASS="LABEL" ID="ROW8" TITLE="Escalation Rules"><input TYPE="CHECKBOX" NAME="ESCALATION_RULES" CHECKED>Escalation Rules</td>
		</tr>
		<tr>
			<td CLASS="LABEL" ID="ROW10" TITLE="Escalation Plan, Escalation Sequence Step"><input TYPE="CHECKBOX" NAME="INFORMATION" CHECKED>Escalation Related</td>
			<td CLASS="LABEL" ID="ROW13" TITLE align="right"><button NAME="Check" STYLE="CURSOR:HAND;WIDTH:80" CLASS="STDBUTTON" ACCESSKEY="C">UnCheck All</button></td>
		</tr>
	</table>
	</div>
	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
		<tr>
			<td colspan="2" HEIGHT="4"></td>
		</tr>
		<tr>
			<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Routing </td>
			<td HEIGHT="5" ALIGN="LEFT">
				<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
					<tr>
						<td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td>
					</tr>
					<tr>
						<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
						<td WIDTH="300" HEIGHT="8"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td>
		</tr>
		<tr>
			<td colspan="2" HEIGHT="1"></td>
		</tr>
	</table>
	<table BORDER="0" WIDTH="100%">
		<tr>
			<td VALIGN="TOP" Width="200">
				<table>
					<tr>
						<td CLASS="LABEL" ID="ROW11" TITLE="Resubmit Reason, Transmission Type, Transmission Type Step, Destination Property, Active Destination(DISABLED)"><input TYPE="CHECKBOX" NAME="COMMON" DISABLED>Common</td>
					</tr>
					<tr>
						<td CLASS="LABEL" ID="ROW16" TITLE="Output Overflow"><input TYPE="CHECKBOX" NAME="OutputOverFlow" CHECKED>Output OverFlow</td>
					</tr>
					<tr>
						<td CLASS="LABEL" ID="ROW12" TITLE="Routing Plan, Transmisison Sequence Step, Output Item, Output Mapping"><input TYPE="CHECKBOX" NAME="CLIENTROUTING">Routing Related</td>
					</tr>
					<tr>
						<td CLASS="LABEL" VALIGN="TOP">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input TYPE="CHECKBOX" NAME="AllRoutingRelatedItems" CHECKED>All Routing Related Items</td>
					</tr>
					<tr>
						<td CLASS="LABEL" ID="ROW15" TITLE="Routing Plan, Transmisison Sequence Step, EDI Outbound Item, EDI Outbound Top Segments, EDI Outbound Segment, EDI Outbound Field, EDI Outbound Mapping"><input TYPE="CHECKBOX" NAME="EDIROUTING" CHECKED>EDI UOF Routing</td>
					</tr>
				</table>
			</td>
			<td VALIGN="TOP">
				<div CLASS="LABEL" ID="SELECTED_RoutingRelatedItems" STYLE="DISPLAY:NONE">
				<table>
					<tr>
						<td CLASS="LABEL">Selected Routing Related Items:<br><OBJECT id=BtnControl_RoutingPlanRelatedItems style="LEFT: 0px; WIDTH: 260px; TOP: 0px; HEIGHT: 23px" type=text/x-scriptlet data=../Scriptlets/ObjButtons.asp?ATTACHCAPTION=Select&amp;REMOVECAPTION=Remove&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDEEDIT=TRUE&amp;HIDENEW=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE>
	</OBJECT></td>
					</tr>
					<tr>
						<td><select CLASS="LABEL" NAME="RoutingRelatedItems_LIST" SIZE="5" STYLE="WIDTH:260;"></select></td>
					</tr>
				</table>
				</div>
			</td>
			<TD VALIGN="TOP" ALIGN="CENTER" CLASS="LABEL">
			<table>
			<tr>
			  <td CLASS="LABEL"><INPUT TYPE="CHECKBOX" NAME="chkLevelOne">Level One Migration</td>
			 </tr>
			 <tr>
			 <td CLASS="LABEL">&nbsp;Reference ID:</td>
			 </tr>
			 <tr>
			 <td>&nbsp;<input TYPE="TEXT" NAME="REFERENCE_ID" CLASS="LABEL" SIZE="50" MAXLENGTH="255"></td>
			 </tr>
			</table>
			</TD>
		</tr>
	</table>
	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
		<tr>
			<td colspan="2" HEIGHT="4"></td>
		</tr>
		<tr>
			<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Output Definition Control Data </td>
			<td HEIGHT="5" ALIGN="LEFT">
				<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
					<tr>
						<td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td>
					</tr>
					<tr>
						<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
						<td WIDTH="300" HEIGHT="8"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td>
		</tr>
		<tr>
			<td colspan="2" HEIGHT="1"></td>
		</tr>
	</table>
	<table BORDER="0" WIDTH="100%">
		<tr>
			<td VALIGN="TOP" Width="200">
				<table>
					<tr>
						<td CLASS="LABEL" ID="ROW14" TITLE="Output Definition, Output Page, Output Field, Output File, Output Subject Body"><input TYPE="CHECKBOX" NAME="OUTPUTDEFINITION">Output Definition</td>
					</tr>
					<tr>
						<td CLASS="LABEL" VALIGN="TOP">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input TYPE="CHECKBOX" NAME="ALLDEFINITIONS" CHECKED>All Definitions</td>
					</tr>
				</table>
			<td VALIGN="TOP">
				<div CLASS="LABEL" ID="SELECTED_OUTPUTDEFS" STYLE="DISPLAY:NONE">
				<table>
					<tr>
						<td CLASS="LABEL">Selected Output Definitions:<br><OBJECT id=BtnControl style="LEFT: 0px; WIDTH: 260px; TOP: 0px; HEIGHT: 23px" type=text/x-scriptlet data=../Scriptlets/ObjButtons.asp?ATTACHCAPTION=Select&amp;REMOVECAPTION=Remove&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDEEDIT=TRUE&amp;HIDENEW=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE>
	</OBJECT></td>
					</tr>
					<tr>
						<td>
							<select CLASS="LABEL" NAME="OUTPUTDEFIDS_LIST" SIZE="5" STYLE="WIDTH:260;"></select>
						</td>
					</tr>
				</table>
				</div>
			</td>
		</tr>
	</table>

	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
		<tr>
			<td colspan="2" HEIGHT="4"></td>
		</tr>
		<tr>
			<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vendor Eligibility </td>
			<td HEIGHT="5" ALIGN="LEFT">
				<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
					<tr>
						<td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td>
					</tr>
					<tr>
						<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
						<td WIDTH="300" HEIGHT="8"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td>
		</tr>
		<tr>
			<td colspan="2" HEIGHT="1"></td>
		</tr>
	</table>

	<table BORDER="0" WIDTH="100%">
		<tr>
			<td VALIGN="TOP" Width="200">
				<table>
					<tr>
						<td CLASS="LABEL" VALIGN="TOP"><input TYPE="CHECKBOX" NAME="VendorDefs" CHECKED>All Definitions</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
		<tr>
			<td colspan="2" HEIGHT="4"></td>
		</tr>
		<tr>
			<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Fraud Detection </td>
			<td HEIGHT="5" ALIGN="LEFT">
				<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
					<tr>
						<td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td>
					</tr>
					<tr>
						<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
						<td WIDTH="300" HEIGHT="8"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td>
		</tr>
		<tr>
			<td colspan="2" HEIGHT="1"></td>
		</tr>
	</table>

	<table BORDER="0" WIDTH="100%" ID="Table3">
		<tr>
			<td VALIGN="TOP" Width="200">
				<table ID="Table4">
					<tr>
						<td CLASS="LABEL" VALIGN="TOP"><input TYPE="CHECKBOX" NAME="FraudDefs" CHECKED>All Definitions</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table5">
		<tr>
			<td colspan="2" HEIGHT="4"></td>
		</tr>
		<tr>
			<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Subrogation </td>
			<td HEIGHT="5" ALIGN="LEFT">
				<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table6">
					<tr>
						<td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td>
					</tr>
					<tr>
						<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
						<td WIDTH="300" HEIGHT="8"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td>
		</tr>
		<tr>
			<td colspan="2" HEIGHT="1"></td>
		</tr>
	</table>

	<table BORDER="0" WIDTH="100%" ID="Table7">
		<tr>
			<td VALIGN="TOP" Width="200">
				<table ID="Table8">
					<tr>
						<td CLASS="LABEL" VALIGN="TOP"><input TYPE="CHECKBOX" NAME="Subrogation" CHECKED ID="Checkbox1">All Definitions</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>


