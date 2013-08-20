<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
Response.Expires=0 
	AccountTextLen = 30	

	Dim SharedCount, SharedCountText, EPID
	SharedCount = 0
	SharedCountText = "Ready"
	
	EPID	= CStr(Request.QueryString("EPID"))

	If EPID <> "" Then
		If EPID = "NEW" Then 
			SharedCount = 0
		End If
	End If	
	
If EPID <> "" Then
	If EPID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = ""
		SQLST = SQLST & "SELECT ESCALATION_PLAN.*, RULES.RULE_TEXT, ACCOUNT_HIERARCHY_STEP.NAME FROM ESCALATION_PLAN, RULES , ACCOUNT_HIERARCHY_STEP WHERE ESCALATION_PLAN_ID = " & EPID & " AND "
		SQLST = SQLST & "ESCALATION_PLAN.ENABLERULE_ID = RULES.RULE_ID(+) AND "
		SQLST = SQLST & "ESCALATION_PLAN.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+)"
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSESCALATION_PLAN_ID = RS("ESCALATION_PLAN_ID")
			RSACCNT_HRCY_STEP_ID = RS("ACCNT_HRCY_STEP_ID")
			RSACCOUNT_NAME = RS("NAME")
			RSLOB_CD = RS("LOB_CD")
			RSRULE_TEXT = RS("RULE_TEXT")
			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
			RSENABLERULE_ID = RS("ENABLERULE_ID")
			RSENABLED_FLG = RS("ENABLED_FLG")
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
End If
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Escalation Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT>
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
var g_StatusInfoAvailable = false;
</SCRIPT>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if EPID <> "" then %>
			<% if SharedCount <= 1 then %>
			document.all.TxtLOB_CD.value = "<%= RSLOB_CD %>"
			document.all.TxtENABLERULE_ID.value = "<%= RSENABLERULE_ID %>"
			document.all.ChkEdit.checked = true
			<% If RSENABLED_FLG="Y" Then %>
				document.all.TxtENABLED_FLG.checked = true
			<% End If %>
			ChkEdit_OnClick
<%	else %>
	SetStatusInfoAvailableFlag(true)
			document.all.ChkEdit.checked = false
			ChkEdit_OnClick
<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>

End Sub

Sub PostTo(strURL)
	FrmDetails.action = "EscalationSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateEPID(inEPID)
	document.all.EPID.value = inEPID
	document.all.spanEPID.innerText = inEPID
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

Function GetEPID
	if document.all.EPID.value <> "NEW" then
		GetEPID = document.all.EPID.value
	else
		GetEPID = ""
	end if 
End Function

'Function GetEPIDName
'	GetEPIDName = document.all.TxtName.value
'End Function

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
errmsg = ""
	If  document.all.TxtLOB_CD.value = "" then
		errmsg = errmsg &  "LOB is a required field."
	end if
	If  document.all.AHSID_ID.innerText = "" then
		errmsg = errmsg &  "AHS ID is a required field."
	end if
	If errmsg = "" Then
		ValidateScreenData = true
	Else
		msgbox errmsg , 0 , "FNSDesigner"
		ValidateScreenData = false
	End If
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.EPID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.EPID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.EPID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.EPID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "ESCALATION_PLAN_ID"& Chr(129) & document.all.EPID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDESCRIPTION.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOB_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ENABLERULE_ID"& Chr(129) & document.all.TxtENABLERULE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ENABLED_FLG"& Chr(129) & Swap(document.all.TxtENABLED_FLG) & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If
	ExeSave = bRet
End Function

Function Swap(Data)
If Data.checked = "False" Then
	Swap = "N"
Else
	Swap = "Y"
End If
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

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"
	
	if document.all.ChkEdit.checked = true then
		SetScreenFieldsReadOnly false,"LABEL"
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScreenFieldsReadOnly true,"DISABLED"
		document.body.setAttribute "ScreenMode", "RO"
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"
end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Function AttachRule (ID, SPANID, strTITLE)
	RID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
			
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ESCALATION&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID, RuleSearchObj.RIDText
	ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID, RuleSearchObj.RIDText
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.value = ""
		SPANID.innerText = ""
	end if
End Function

Function DetachAccount(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

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
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ESCALATION&SELECTONLY=TRUE&AHSID=" &AHSID
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
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Escalation Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="EscalationSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchEPID" value="<%=Request.QueryString("SearchEPID")%>">
<input type="hidden" name="SearchACCNT_HRCY_STEP_ID" value="<%=Request.QueryString("SearchACCNT_HRCY_STEP_ID")%>">
<input type="hidden" name="SearchLOB_CD" value="<%=Request.QueryString("SearchLOB_CD")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="EPID" value="<%=Request.QueryString("EPID")%>" >

<%	

If EPID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</td>
<td>
<input  ScrnBtn="TRUE"  TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>

<table class="LABEL">
<tr>
	<td>
	<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::DetachAccount AHSID_ID, AHSID_TEXT">
	</td>
	<td width=305 nowrap>Account:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(RSACCOUNT_NAME)%>" ><%=TruncateText(RSACCOUNT_NAME,AccountTextLen)%></SPAN></td>
	<td>A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL><%=RSACCNT_HRCY_STEP_ID%></SPAN></td>
	</tr>
</table>
<table CLASS="LABEL" CELLPADDING=0 CELLSPACING=0 >
<tr>
<td>
<table class="LABEL">
	<tr>
	<td CLASS=LABEL COLSPAN=2>Escalation ID:&nbsp<span id="spanEPID"><%=Request.QueryString("EPID")%></span></td>
	</tr>
	<tr>
	<td CLASS="LABEL">LOB:<br><select ScrnBtn="TRUE" NAME="TxtLOB_CD" CLASS="LABEL" tabindex=3><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD","",true)%></select></td>
	</tr>
	</TABLE>
	<TABLE>
	<tr>
	<td CLASS=LABEL COLSPAN=4>Description:<br><input ScrnInput="TRUE" size=80 CLASS="LABEL" MAXLENGTH=80 TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</TR>
	<TR>
	<td nowrap CLASS=LABEL>Enabled Rule:<BR>
	<INPUT TYPE=Text SIZE=10 READONLY CLASS="LABEL" STYLE="BACKGROUND-COLOR:SILVER" NAME=TxtENABLERULE_ID VALUE="<%= RSENABLERULE_ID %>">
	<IMG NAME=BtnAttachValid STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule TxtENABLERULE_ID, ENABLED_TEXT,'Valid'">
	<IMG NAME=BtnDetachValid STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach TxtENABLERULE_ID, ENABLED_TEXT">
	<SPAN ID=ENABLED_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(RSRULE_TEXT)%>" ><%=TruncateText(RSRULE_TEXT,AccountTextLen)%></SPAN>
	</TD></TR>
	<TR>
	<TD CLASS=LABEL><input ScrnBtn="TRUE" CLASS="LABEL" TYPE="CHECKBOX" NAME="TxtENABLED_FLG" ONCLICK="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	Enabled:</td>
	</TR>
	</TABLE>
<% Else %>
<DIV style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Escalation selected.
</DIV>
<% End If %>
</form>
</body>
</html>


