<%
'***************************************************************
'form to enter/edit Mailbox Assignment Rules
'
'$History: MailboxAssignTypeDetails.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:47p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->


<%	Response.Expires = 0 
	Response.AddHeader  "Pragma", "no-cache"
	Response.Buffer = true
	RuleTextLen = 30
	RSAHSID = Request.QueryString("AHSID")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Mailbox Assignment Type Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CMailboxAssignRuleSearchObj()
{
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
function CBranchSearchObj()
{
	this.BID = "";
	this.BNumber = "";
	this.Selected = false;
}

var AHSSearchObj = new CAHSSearchObj();
var RuleSearchObj = new CRuleSearchObj();
var MailboxAssignRuleSearchObj = new CMailboxAssignRuleSearchObj();
var BranchSearchObj = new CBranchSearchObj();
var g_StatusInfoAvailable = false;

</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 230;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
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
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_MAILBOX_ASSIGNMENT&RID=" & RID
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

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText(SPANID,inText)
	If Len(inText) < <%=RuleTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid (inText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Function AttachBranch(ID)
	BID = cBranchID
	MODE = document.body.getAttribute("ScreenMode")

	BranchSearchObj.BID = BID
	BranchSearchObj.Selected = false
	If BID = "" Then BID = "NEW"
	If BID = "NEW" And MODE = "RO" Then
		MsgBox "No Branch currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Branch\BranchMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_MAILBOX_ASSIGNMENT&BID=" & BID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,BranchSearchObj ,"dialogWidth:550px;dialogHeight:450px;center"
	If BranchSearchObj.Selected Then
		If BranchSearchObj.BNum <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = BranchSearchObj.BNum
			cBranchID = BranchSearchObj.BID
		end if
	End If

End Function

Function DetachBranch (ID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
	end if
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
	
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_BRANCH_ASSIGNMENT&SELECTONLY=TRUE&AHSID=" &AHSID
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


Sub PostTo(strURL)
	FrmDetails.action = "MailboxAssignTypeSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateMATID(inMATID)
	document.all.MATID.value = inMATID
	document.all.spanMATID.innerText = inMATID
	'	A MAILBOX_ASSIGNMENT_RULE is required
	Dim MATID, MARID, MODE
	MARID = "NEW"
	MATID = document.all.MATID.value
	MODE = document.body.getAttribute("ScreenMode")
	
	MailboxAssignRuleSearchObj.Selected = false

	strURL = "MailboxAssignRuleModal.asp?MATID=" & inMATID & "&MARID=NEW" & "&MODE=" & MODE & "&RequiredMsg=Y"
	showModalDialog  strURL,MailboxAssignRuleSearchObj ,"center"
	If MailboxAssignRuleSearchObj.Selected = true Then Refresh
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

Function GetMATID
	if document.all.MATID.value <> "NEW" then
		GetMATID = document.all.MATID.value
	else
		GetMATID = ""
	end if 
End Function

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Function f_CheckIsThisRequired
	IF CStr(document.all.getAttribute("IsThisRequired")) = "Y" Then
		f_CheckIsThisRequired = true
	ELSE
		f_CheckIsThisRequired = False
	END IF
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	If  document.all.TxtDescription.value = "" then
		MsgBox "Description is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.AHSID_ID.innerText = "" then
		MsgBox "A.H. Step ID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.BRANCH_NUMBER.innerText = "" then
		MsgBox "Branch Number is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	ValidateScreenData = true
End Function

Function GetSelectedMARID
	GetSelectedMARID = document.frames("DataFrame").GetSelectedMARID
End Function

Function f_IsThisLastRecord
	f_IsThisLastRecord = document.frames("DataFrame").f_LastBARuleRecord()
End Function

Sub ExeNewMailboxRule
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.MATID.value = "" Or document.all.MATID.value = "NEW" Then
		Exit Sub
	End If

<%If HasAddPrivilege("FNSD_MAILBOX_ASSIGNMENT","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to add Mailbox assignment rules.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		


	dim MATID, MARID, MODE
	MARID = "NEW"
	MATID = document.all.MATID.value
	MODE = document.body.getAttribute("ScreenMode")
	
	MailboxAssignRuleSearchObj.Selected = false

	strURL = "MailboxAssignRuleModal.asp?MATID=" & MATID & "&MARID=" & MARID & "&MODE=" & MODE 	
	showModalDialog  strURL,MailboxAssignRuleSearchObj ,"center"

	If MailboxAssignRuleSearchObj.Selected = true Then	Refresh
End Sub

Sub Refresh
	MATID = document.all.MATID.value
	document.all.tags("IFRAME").item("DataFrame").src = "MailboxAssignTypeDetailsData.asp?MATID=" & MATID
End Sub

Sub ExeEditMailboxRule
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.MATID.value = "" Or document.all.MATID.value = "NEW" Then
		Exit Sub
	End If
	
<%If HasModifyPrivilege("FNSD_MAILBOX_ASSIGNMENT","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to edit Mailbox assignment rules.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

	dim MARID, MATID
	MARID = GetSelectedMARID
	MATID = document.all.MATID.value
	
	If MARID <> "" Then
		MailboxAssignRuleSearchObj.Selected = false
		strURL = "MailboxAssignRuleModal.asp?MATID=" & MATID & "&MARID=" & MARID & "&MODE=" & MODE 	
		showModalDialog  strURL,MailboxAssignRuleSearchObj ,"center"
		If MailboxAssignRuleSearchObj.Selected = true Then	Refresh
	Else
		MsgBox "Please select a Mailbox assignment rule to Edit.", 0, "FNSNet Designer"		
	End If
	
End Sub

Sub ExeRemoveMailboxRule
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.MATID.value = "" Or document.all.MATID.value = "NEW" Then
		Exit Sub
	End If

	If f_IsThisLastRecord Then
		Msgbox "MAILBOX_ASSIGNMENT_TYPE must have at least 1 MAILBOX_ASSIGNMENT_RULE.  Delete Failed.",0,"FNSDesigner"
		Exit Sub
	End If
<%If HasDeletePrivilege("FNSD_MAILBOX_ASSIGNMENT","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to delete Mailbox assignment rules.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

	dim MARID, sResult
	MARID = GetSelectedMARID
	
	If MARID <> "" Then
		sResult = sResult &  MARID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmDetails.action = "MailboxAssignRuleSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a Mailbox assignment rule to Remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub

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
	
	If document.all.MATID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
	
	FrmDetails.action = "MailboxAssignTypeCopy.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit
	ExeCopy = true
End Function

Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.MATID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	if ValidateScreenData = false then 
		ExeSave = false
		exit function
	end if

	If document.all.MATID.value = "NEW" then
		document.all.TxtAction.value = "INSERT"
	else
		document.all.TxtAction.value = "UPDATE"
	end if
	sResult = sResult & "MAILBOX_ASSIGNMENT_TYPE_ID"& Chr(129) & document.all.MATID.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
	sResult = sResult & "RULE_ID"& Chr(129) & document.all.RULE_ID.innerText & Chr(129) & "1" & Chr(128)
	sResult = sResult & "BRANCH_NUMBER"& Chr(129) & document.all.BRANCH_NUMBER.innerText & Chr(129) & "1" & Chr(128)
	
	document.all.TxtSaveData.Value = sResult
	FrmDetails.action = "MailboxAssignTypeSave.asp"
	FrmDetails.method = "POST"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit
	bRet = true
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
<!--#include file="..\lib\MABtnControl.inc"-->

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Mailbox Assignment Type Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="MailboxAssignTypeSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchMATID" value="<%=Request.QueryString("SearchMATID")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchRuleID" value="<%=Request.QueryString("SearchRuleID")%>">
<input type="hidden" name="SearchRuleText" value="<%=Request.QueryString("SearchRuleText")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="MATID" value="<%=Request.QueryString("MATID")%>">

<%	
Dim MATID
MATID	= CStr(Request.QueryString("MATID"))
If MATID <> "" Then
	If MATID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT MAILBOX_ASSIGNMENT_TYPE.*, RULES.*, ACCOUNT_HIERARCHY_STEP.NAME NAME FROM " &_
				"MAILBOX_ASSIGNMENT_TYPE, RULES, ACCOUNT_HIERARCHY_STEP WHERE MAILBOX_ASSIGNMENT_TYPE.RULE_ID = RULES.RULE_ID(+) AND " &_
				"MAILBOX_ASSIGNMENT_TYPE.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				"MAILBOX_ASSIGNMENT_TYPE_ID = " & MATID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
			RSAHSID_TEXT = ReplaceQuotesInText(RS("NAME"))
			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
			RSRULE_ID = RS("RULE_ID")			
			RSRULE_TEXT = ReplaceQuotesInText(RS("RULE_TEXT"))
			RSBRANCH_NUM = RS("BRANCH_NUMBER")			
		End If
		RS.Close
		'	get the barnch_id
		if len(RSBRANCH_NUM) <> 0 then
			Set RS = Conn.Execute("Select BRANCH_ID From BRANCH Where BRANCH_NUMBER='" & RSBRANCH_NUM & "'")
			If Not RS.EOF Then
				RSBRANCH_ID = RS("BRANCH_ID")			
			end if
			RS.Close
		end if
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>
<script LANGUAGE="vbscript">
dim cBranchID

cBranchID = "<%=RSBRANCH_ID%>"
</script>
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
<tr><td colspan=2>Mailbox Assignment Type ID:&nbsp;<span id="spanMATID"><%=Request.QueryString("MATID")%></span></td></tr>
<tr>
	<td colspan=2>Description:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="65" TYPE="TEXT" NAME="TxtDescription" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>

<table class="Label">
<td>
<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
</td>
<td width=305 nowrap>Account:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=RSAHSID_TEXT%>"><%=TruncateText(RSAHSID_TEXT,RuleTextLen)%></SPAN></td>
<td>A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL><%=RSAHSID%></SPAN><input name=TxtAHSID type=hidden value="<%=RSAHSID%>"></input></td>
</table>

<table class="Label">
<td>
<IMG NAME=BtnAttachRule STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
<IMG NAME=BtnDetachRule STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
</td>
<td width=305 nowrap>Rule Text:&nbsp;<SPAN ID=RULE_TEXT CLASS=LABEL TITLE="<%=RSRULE_TEXT%>" ><%=TruncateText(RSRULE_TEXT,RuleTextLen)%></SPAN></td>
<td>Rule ID:&nbsp;<SPAN ID=RULE_ID CLASS=LABEL><%=RSRULE_ID%></SPAN></td>
</table>

<table class="Label" ID="Table2">
<td>
<IMG NAME=BtnAttachBranch STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachBranch BRANCH_NUMBER">
<IMG NAME=BtnDetachBranch STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachBranch BRANCH_NUMBER">
</td>
<td width=305 nowrap>Branch Number:&nbsp;<SPAN ID="BRANCH_NUMBER" CLASS=LABEL><%=RSBRANCH_NUM%></SPAN></td>
</table>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Mailbox Assignment Rules</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<span class="Label" ID=SPANDATA>Retrieving...</span>
<fieldset id="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<OBJECT id=MABtnControl style="LEFT: 0px; WIDTH: 100%; HEIGHT: 23px" type=text/x-scriptlet data=../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&amp;HIDEATTACH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE VIEWASTEXT>
	</OBJECT>
<iframe width=100% height=100% name="DataFrame" src="MailboxAssignTypeDetailsData.asp?<%=Request.QueryString%>"></iframe> 
</fieldset>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Mailbox assignment type selected.
</div>


<% End If %>

</form>
</body>
</html>


