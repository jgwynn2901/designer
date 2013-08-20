<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires = 0 
	Response.Buffer = true
	RuleTextLen = 30
	RoutingTextLen = 60 

	RSAHSID = Request.QueryString("AHSID")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Routing Address Rule Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CRoutingAddressSearchObj()
{
	this.RAID = "";
	this.RAIDDescription = "";
	this.RAIDState = "";
	this.RAIDFIPS = "";
	this.RAIDZip = "";
	this.Selected = false;
}
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();
var RuleSearchObj = new CRuleSearchObj();
var RoutingAddressSearchObj = new CRoutingAddressSearchObj();
var g_StatusInfoAvailable = false;

</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Function AttachRule (ID, SPANID, strTITLE)
	RID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
		If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ROUTING_ADDRESS_RULE&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID,RuleSearchObj.RIDText
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
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ROUTING_ADDRESS_RULE&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,AHSSearchObj ,"center"

	'if Selected=true update everything, otherwise if AHSID is the same, update text in case of save
	If AHSSearchObj.Selected = true Then
		If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = AHSSearchObj.AHSID
		end if
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=RuleTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub


Function AttachRoutingAddress
	RAID = document.all.spanRAID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	RoutingAddressSearchObj.RAID = RAID
	RoutingAddressSearchObj.RAIDDescription = spanDESCRIPTION.innerText
	RoutingAddressSearchObj.RAIDState = spanSTATE.innerText
	RoutingAddressSearchObj.RAIDFIPS = spanFIPS.innerText
	RoutingAddressSearchObj.RAIDZip = spanZip.innerText
	RoutingAddressSearchObj.Selected = false

	If RAID = "" Then RAID = "NEW"

	If RAID = "NEW" And MODE = "RO" Then
		MsgBox "No routing address currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If

	strURL = "RoutingAddressMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ROUTING_ADDRESS_RULE&RAID=" & RAID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RoutingAddressSearchObj ,"center"

	'if Selected=true update everything, otherwise if RAID is the same, update text in case of save
	If RoutingAddressSearchObj.Selected = true Then
		If RoutingAddressSearchObj.RAID <> document.all.spanRAID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			document.all.spanRAID.innerText = RoutingAddressSearchObj.RAID
		end if
		UpdateRoutingAddressFields
	ElseIf document.all.spanRAID.innerText = RoutingAddressSearchObj.RAID And RoutingAddressSearchObj.RAID<> "" Then
		UpdateRoutingAddressFields
	End If

End Function


Function DetachRoutingAddress
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		spanRAID.innerText = ""
		spanDESCRIPTION.innerText = ""
		spanSTATE.innerText = ""
		spanFIPS.innerText = ""
		spanZIP.innerText = ""
	end if
End Function

Sub UpdateRoutingAddressFields
	If Len(RoutingAddressSearchObj.RAIDDescription) < <%=RoutingTextLen%> Then
		spanDESCRIPTION.innertext = RoutingAddressSearchObj.RAIDDescription
	Else
		spanDESCRIPTION.innertext = Mid ( RoutingAddressSearchObj.RAIDDescription, 1, <%=RoutingTextLen%>) & " ..."
	End If
	spanDESCRIPTION.title = RoutingAddressSearchObj.RAIDDescription

	If Len(RoutingAddressSearchObj.RAIDState) < <%=RoutingTextLen%> Then
		spanSTATE.innertext = RoutingAddressSearchObj.RAIDState
	Else
		spanSTATE.innertext = Mid ( RoutingAddressSearchObj.RAIDState, 1, <%=RoutingTextLen%>) & " ..."
	End If
	spanSTATE.title = RoutingAddressSearchObj.RAIDState

	If Len(RoutingAddressSearchObj.RAIDFIPS) < <%=RoutingTextLen%> Then
		spanFIPS.innertext = RoutingAddressSearchObj.RAIDFIPS
	Else
		spanFIPS.innertext = Mid ( RoutingAddressSearchObj.RAIDFIPS, 1, <%=RoutingTextLen%>) & " ..."
	End If
	spanFIPS.title = RoutingAddressSearchObj.RAIDFIPS

	If Len(RoutingAddressSearchObj.RAIDZip) < <%=RoutingTextLen%> Then
		spanZIP.innertext = RoutingAddressSearchObj.RAIDZip
	Else
		spanZIP.innertext = Mid ( RoutingAddressSearchObj.RAIDZip, 1, <%=RoutingTextLen%>) & " ..."
	End If
	spanZIP.title = RoutingAddressSearchObj.RAIDZip

End Sub


Sub PostTo(strURL)
	FrmDetails.action = "RoutingAddressRuleSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateRARID(inRARID)
	document.all.RARID.value = inRARID
	document.all.spanRARID.innerText = inRARID
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

Function GetRARID
	if document.all.RARID.value <> "NEW" then
		GetRARID = document.all.RARID.value
	else
		GetRARID = ""
	end if 
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
	If  document.all.spanRAID.innerText = "" then
		MsgBox "Routing Address is required.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.TxtLOBCD.value = "" then
		MsgBox "LOB is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.AHSID_ID.innerText = "" then
		MsgBox "A.H. Step ID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	ValidateScreenData = true
End Function

Function InEditMode
	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		InEditMode = false
	End If
End Function

Function ExeCopy
	If Not InEditMode Then
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.RARID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
	
	document.body.setAttribute "ScreenDirty","YES"
	document.all.RARID.value = "NEW"
	ExeCopy = ExeSave
End Function


Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.RARID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.RARID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "ROUTINGADDRESSRULE_ID"& Chr(129) & document.all.RARID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOBCD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTINGRULE_ID"& Chr(129) & document.all.RULE_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTINGADDRESS_ID"& Chr(129) & document.all.spanRAID.innerText & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "RoutingAddressRuleSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
			
		bRet = true
		
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If
	
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

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Routing Address Rule Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="RoutingAddressRuleSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchRARID" value="<%=Request.QueryString("SearchRARID")%>">
<input type="hidden" name="SearchLOBCD" value="<%=Request.QueryString("SearchLOBCD")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchRuleID" value="<%=Request.QueryString("SearchRuleID")%>">
<input type="hidden" name="SearchRuleText" value="<%=Request.QueryString("SearchRuleText")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="RARID" value="<%=Request.QueryString("RARID")%>">

<%	
Dim RARID
RARID	= CStr(Request.QueryString("RARID"))
If RARID <> "" Then
	If RARID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT ROUTINGADDRESSRULE.*,ROUTINGADDRESS.*, RULES.*,  ACCOUNT_HIERARCHY_STEP.NAME FROM ROUTINGADDRESSRULE,ROUTINGADDRESS, RULES, ACCOUNT_HIERARCHY_STEP WHERE " &_
				"ROUTINGADDRESSRULE.ROUTINGRULE_ID = RULES.RULE_ID(+) AND " &_
				"ROUTINGADDRESSRULE.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				"ROUTINGADDRESSRULE.ROUTINGADDRESS_ID = ROUTINGADDRESS.ROUTINGADDRESS_ID(+) AND " &_
				"ROUTINGADDRESSRULE_ID = " & RARID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
			RSAHSID_TEXT = ReplaceQuotesInText(RS("NAME"))
			RSLOBCD = RS("LOB_CD")
			RSRULE_ID = RS("ROUTINGRULE_ID")			
			RSRULE_TEXT = ReplaceQuotesInText(RS("RULE_TEXT"))

			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
			RSSTATE = ReplaceQuotesInText(RS("STATE"))
			RSFIPS = ReplaceQuotesInText(RS("FIPS"))
			RSZIP = ReplaceQuotesInText(RS("ZIP"))
			RSROUTINGADDRESS_ID = RS("ROUTINGADDRESS_ID")

		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" >
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td colspan=2>Routing Address Rule ID:&nbsp;<span id="spanRARID"><%=Request.QueryString("RARID")%></span></td></tr>
<tr>
	<td>LOB:<br><select ScrnBtn="TRUE" name="TxtLOBCD" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOBCD,true)%></select></td>
</tr>
</table>

<table class="LABEL">
<tr>
	<td>
	<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	</td>
	<td width=305 nowrap>Account:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=RSAHSID_TEXT%>" ><%=TruncateText(RSAHSID_TEXT,RuleTextLen)%></SPAN></td>
	<td>A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL><%=RSAHSID%></SPAN></td>
	</tr>
</table>


<table class="Label">
<td>
<IMG NAME=BtnAttachRule STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
<IMG NAME=BtnDetachRule STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
</td>
<td width=305 nowrap>Rule Text:&nbsp;<SPAN ID=RULE_TEXT CLASS=LABEL TITLE="<%=RSRULE_TEXT%>" ><%=TruncateText(RSRULE_TEXT,RuleTextLen)%></SPAN></td>
<td>Rule ID:&nbsp;<SPAN ID=RULE_ID CLASS=LABEL><%=RSRULE_ID%></SPAN></td>
</table>
<br>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Routing Address</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table class="Label">
<td>
<IMG NAME=BtnAttachAddress STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Routing Address" ONCLICK="VBScript::AttachRoutingAddress">
<IMG NAME=BtnDetachAddress STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Routing Address" OnClick="VBScript::DetachRoutingAddress">
</td>
<td>Routing Address ID:&nbsp;<SPAN ID=spanRAID CLASS=LABEL><%=RSROUTINGADDRESS_ID%></SPAN></td>
</table>

<table>
<tr NOWRAP CLASS="LABEL"><td COLSPAN=1><td>Description:<td><SPAN ID=spanDESCRIPTION CLASS=LABEL TITLE="<%=RSDESCRIPTION%>"><%=TruncateText(RSDESCRIPTION,RoutingTextLen)%></SPAN></tr>
<tr NOWRAP CLASS="LABEL"><td COLSPAN=1><td>State:<td><SPAN ID=spanSTATE CLASS=LABEL TITLE="<%=RSSTATE%>"><%=TruncateText(RSSTATE,RoutingTextLen)%></SPAN></tr>
<tr NOWRAP CLASS="LABEL"><td COLSPAN=1><td>FIPS:<td><SPAN ID=spanFIPS CLASS=LABEL TITLE="<%=RSFIPS%>"><%=TruncateText(RSFIPS,RoutingTextLen)%></SPAN></tr>
<tr NOWRAP CLASS="LABEL"><td COLSPAN=1><td>Zip:<td><SPAN ID=spanZIP CLASS=LABEL TITLE="<%=RSZIP%>"><%=TruncateText(RSZIP,RoutingTextLen)%></SPAN></tr>
</table>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No routing address rule selected.
</div>


<% End If %>

</form>
</body>
</html>


