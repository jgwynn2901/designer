<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires = 0 
	Response.Buffer = true
	RuleTextLen = 30
	BranchTextLen = 30

	RSAHSID = Request.QueryString("AHSID")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Claim Number Assignment Rule Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CBranchSearchObj()
{
	this.BID = "";
	this.BIDOfficeName = "";
	this.Selected = false;
}
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}

var AHSSearchObj = new CAHSSearchObj();
var BranchSearchObj = new CBranchSearchObj();
var RuleSearchObj = new CRuleSearchObj();

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
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CLAIM_ASSIGNMENT&RID=" & RID
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
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CLAIM_ASSIGNMENT&SELECTONLY=TRUE&AHSID=" &AHSID
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

Function AttachBranch (ID, SPANID, strTITLE)
	BID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	BranchSearchObj.BID = BID
	BranchSearchObj.BIDOfficeName = SPANID.title
	BranchSearchObj.Selected = false

	If BID = "" Then BID = "NEW"
	
	If BID = "NEW" And MODE = "RO" Then
		MsgBox "No branch currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Branch\BranchMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CLAIM_ASSIGNMENT&SELECTONLY=TRUE&BranchTypeFilter=CLAIMHANDLING&BID=" & BID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,BranchSearchObj ,"center"

	'if Selected=true update everything, otherwise if BID is the same, update text in case of save
	If BranchSearchObj.Selected = true Then
		If BranchSearchObj.BID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = BranchSearchObj.BID
		end if
		UpdateBranchText(SPANID)
	ElseIf ID.innerText = BranchSearchObj.BID And BranchSearchObj.BID<> "" Then
		UpdateBranchText(SPANID)
	End If

End Function

Sub UpdateBranchText (SPANID)
	If Len(BranchSearchObj.BIDOfficeName) < <%=BranchTextLen%> Then
		SPANID.innertext = BranchSearchObj.BIDOfficeName
	Else
		SPANID.innertext = Mid ( BranchSearchObj.BIDOfficeName, 1, <%=BranchTextLen%>) & " ..."
	End If
	SPANID.title = BranchSearchObj.BIDOfficeName
End Sub


Sub PostTo(strURL)
	FrmDetails.action = "ClaimAssignRuleSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateCARID(inCARID)
	document.all.CARID.value = inCARID
	document.all.spanCARID.innerText = inCARID
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

Function GetCARID
	if document.all.CARID.value <> "NEW" then
		GetCARID = document.all.CARID.value
	else
		GetCARID = ""
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
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.CARID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.CARID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.CARID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		sResult = sResult & "CLAIMNUMBERASSIGNMENTRULE_ID"& Chr(129) & document.all.CARID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOBCD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BRANCH_ID"& Chr(129) & document.all.BRANCH_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "RULE_ID"& Chr(129) & document.all.RULE_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQUENCE"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LENGTH"& Chr(129) & document.all.TxtLength.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PREFIX"& Chr(129) & document.all.TxtPrefix.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SUFFIX"& Chr(129) & document.all.TxtSuffix.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "WARNINGPERCENT"& Chr(129) & document.all.TxtWarningPercent.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MINIMUM"& Chr(129) & document.all.TxtMinimum.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAXIMUM"& Chr(129) & document.all.TxtMaximum.value & Chr(129) & "1" & Chr(128)
		if document.all.ChkReuseFlg.checked = True then
			sResult = sResult & "REUSE_FLG"& Chr(129) & "Y"  & Chr(129) & "1" & Chr(128)
		else 
			sResult = sResult & "REUSE_FLG"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if
		sResult = sResult & "ASSIGN_TO_ATTR_NAME"& Chr(129) & UCase(Trim(document.all.TxtAttrName.value)) & Chr(129) & "1" & Chr(128)

		If document.all.CARID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
			if document.all.TxtMinimum.value <> "" then
				sResult = sResult & "NEXT" & Chr(129) & document.all.TxtMinimum.value & Chr(129) & "1" & Chr(128)
			end if
		else
			document.all.TxtAction.value = "UPDATE"
		end if
	
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "ClaimAssignRuleSave.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Claim Number Assignment Rule Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="ClaimAssignRuleSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchCARID" value="<%=Request.QueryString("SearchCARID")%>">
<input type="hidden" name="SearchBID" value="<%=Request.QueryString("SearchBID")%>">
<input type="hidden" name="SearchLOBCD" value="<%=Request.QueryString("SearchLOBCD")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchRuleID" value="<%=Request.QueryString("SearchRuleID")%>">
<input type="hidden" name="SearchRuleText" value="<%=Request.QueryString("SearchRuleText")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="SearchAttrName" value="<%=Request.QueryString("SearchAttrName")%>"">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="CARID" value="<%=Request.QueryString("CARID")%>">

<%	
Dim CARID
CARID	= CStr(Request.QueryString("CARID"))
If CARID <> "" Then
	If CARID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT CLAIMNUMBERASSIGNMENTRULE.*, RULES.RULE_TEXT, BRANCH.OFFICE_NAME, ACCOUNT_HIERARCHY_STEP.NAME FROM " &_
				"CLAIMNUMBERASSIGNMENTRULE, RULES, BRANCH, ACCOUNT_HIERARCHY_STEP WHERE " &_
				"CLAIMNUMBERASSIGNMENTRULE.RULE_ID = RULES.RULE_ID(+) AND " &_
				"CLAIMNUMBERASSIGNMENTRULE.BRANCH_ID = BRANCH.BRANCH_ID(+) AND " &_
				"CLAIMNUMBERASSIGNMENTRULE.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				"CLAIMNUMBERASSIGNMENTRULE_ID = " & CARID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
			RSAHSID_TEXT = ReplaceQuotesInText(RS("NAME"))
			RSLOBCD = RS("LOB_CD")
			RSRULE_ID = RS("RULE_ID")			
			RSRULE_TEXT = ReplaceQuotesInText(RS("RULE_TEXT"))
			RSBRANCH_ID = RS("BRANCH_ID")			
			RSBRANCH_OFFICE_NAME = ReplaceQuotesInText(RS("OFFICE_NAME"))
			RSSEQUENCE = RS("SEQUENCE")
			RSLENGTH = RS("LENGTH")
			RSPREFIX = ReplaceQuotesInText(RS("PREFIX"))
			RSSUFFIX = ReplaceQuotesInText(RS("SUFFIX"))
			RSMINIMUM = RS("MINIMUM")
			RSMAXIMUM = RS("MAXIMUM")
			RSWARNINGPERCENT = RS("WARNINGPERCENT")
			RSREUSE_FLG = RS("REUSE_FLG")
			RSATTRNAME = RS("ASSIGN_TO_ATTR_NAME")
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If

	SELECTEDLOB = ""

	If Not IsNull(RSLOBCD) Then
		temp = Trim(CStr(RSLOBCD))
		If  temp <> "" And temp = "*" Then	SELECTEDLOB = "SELECTED"
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
<tr><td colspan=2>Claim Number Assignment Rule ID:&nbsp;<span id="spanCARID"><%=Request.QueryString("CARID")%></span></td></tr>
<tr>
	<td>LOB:<br><select ScrnBtn="TRUE" name="TxtLOBCD" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOBCD,true)%><option value="*  " <%=SELECTEDLOB%> >*</option></select></td>
</tr>
</table> 

<table class="Label" ONDRAGSTART="return false;">
<tr>
<td>
<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
</td>
<td width=305 nowrap>Account:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(RSAHSID_TEXT)%>" ><%=TruncateText(RSAHSID_TEXT,RuleTextLen)%></SPAN></td>
<td>A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL><%=RSAHSID%></SPAN></td>
</tr>
<tr>
	<td>
	<IMG NAME=BtnAttachRule STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
	<IMG NAME=BtnDetachRule STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
	</td>
	<td nowrap WIDTH=300>Rule Text:&nbsp<SPAN ID=RULE_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(RSRULE_TEXT)%>" ><%=TruncateText(RSRULE_TEXT,RuleTextLen)%></SPAN></td>
	<td>Rule ID:&nbsp<span ID=RULE_ID><%=RSRULE_ID%></span></td>
</tr>	
<tr>
	<td>
	<IMG NAME=BtnAttachBranch STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Branch" ONCLICK="VBScript::AttachBranch BRANCH_ID, BRANCH_DESC,''">
	<IMG NAME=BtnDetachBranch STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Branch" OnClick="VBScript::Detach BRANCH_ID, BRANCH_DESC">
	</td>
	<td nowrap  WIDTH=300>Branch:&nbsp<SPAN ID=BRANCH_DESC CLASS=LABEL TITLE="<%=RSBRANCH_OFFICE_NAME%>" ><%=TruncateText(RSBRANCH_OFFICE_NAME,BranchTextLen)%></SPAN></td>
	<td>Branch ID:<span ID=BRANCH_ID><%=RSBRANCH_ID%></span></td>
</tr>


</table>


<table CLASS="LABEL" >
<tr>
	<td>Sequence<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Length<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtLength" VALUE="<%=RSLENGTH%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Prefix<br><input ScrnInput="TRUE" MAXLENGTH="70" CLASS="LABEL" size="30" TYPE="TEXT" NAME="TxtPrefix" VALUE="<%=RSPREFIX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Suffix<br><input ScrnInput="TRUE" MAXLENGTH="70" CLASS="LABEL" size="30" TYPE="TEXT" NAME="TxtSuffix" VALUE="<%=RSSUFFIX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
</tr>
<tr>
	<td>Minimum<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtMinimum" VALUE="<%=RSMINIMUM%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Maximum<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtMaximum" VALUE="<%=RSMAXIMUM%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Warning Percent<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtWarningPercent" VALUE="<%=RSWARNINGPERCENT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Reuse?<input ScrnBtn="TRUE" CLASS="LABEL" TYPE="CHECKBOX" NAME="ChkReuseFlg"  <% If CStr(RSREUSE_FLG) = "Y" Then Response.Write("CHECKED")%> ONCLICK="VBScript::Control_OnChange"></input></td>
</tr>
</table>
<table CLASS="LABEL"> 
<tr>
	<td>Assign-To Attribute Name (* Optional)<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="44" TYPE="TEXT" NAME="TxtAttrName" VALUE="<%=RSATTRNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
</tr>
</table>


<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Claim Number Assignment Rule selected.
</div>


<% End If %>

</form>
</body>
</html>


