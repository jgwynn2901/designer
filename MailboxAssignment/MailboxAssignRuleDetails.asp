<%
'***************************************************************
'displays Mailbox Assignment Rule Details.
'
'$History: MailboxAssignRuleDetails.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:46p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'*****************  Version 1  *****************
'User: Roberto.agit    Date:  4/05/04  Time: 11:26a
'Created MailboxAssignRuleDetails.asp
'Comment:
  
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim MATID, MARID, isRequired
	
	MATID =  CStr(Request.QueryString("MATID"))
	MARID =  CStr(Request.QueryString("MARID"))
	isRequired = Request.QueryString("RequiredMsg")
	
	IF isRequired = "Y" Then
		s_DisplayMsg = "At least 1 MAILBOX_ASSIGNMENT_RULE is required"
	Else
		s_DisplayMsg = "Ready"
	End If
	
	MailboxTextLen = 30
	RuleTextLen = 30
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Mailbox Assignment Rule Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CMailboxSearchObj()
{
	this.MBID = "";
	this.Selected = false;
}
var MailboxSearchObj = new CMailboxSearchObj();
var RuleSearchObj = new CRuleSearchObj();

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

</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

<!--#include file="..\lib\Help.asp"-->

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
End Sub

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
		UpdateRuleText(SPANID)
	ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateRuleText(SPANID)
	End If

End Function


Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateRuleText (SPANID)
	If Len(RuleSearchObj.RIDText) < <%=RuleTextLen%> Then
		SPANID.innertext = RuleSearchObj.RIDText
	Else
		SPANID.innertext = Mid ( RuleSearchObj.RIDText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = RuleSearchObj.RIDText
End Sub

Function AttachMailbox (ID)
	MBID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	MailboxSearchObj.MBID = MBID
	MailboxSearchObj.Selected = false

	If MBID = "" Then MBID = "NEW"
	
	If MBID = "NEW" And MODE = "RO" Then
		MsgBox "No Mailbox currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Mailbox\MailboxMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_MAILBOX_ASSIGNMENT&SELECTONLY=TRUE&MBID=" & MBID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,MailboxSearchObj ,"center"

	'if Selected=true update everything, otherwise if BID is the same, update text in case of save
	If MailboxSearchObj.Selected = true Then
		If MailboxSearchObj.MBID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = MailboxSearchObj.MBID
		end if
	End If

End Function

Function DetachMailbox(ID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
	end if
End Function

Sub UpdateMARID(inMARID)
	document.all.MARID.value = inMARID
	document.all.spanMARID.innerText = inMARID
	if document.all.spanMARID.innerText <> "NEW" then
			document.body.setAttribute "IsThisRequired", "N"
	End If
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

Function GetMARID
	if document.all.MARID.value <> "NEW" then
		GetMARID = document.all.MARID.value
	else
		GetMARID = ""
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
	If CStr(document.body.getAttribute("IsThisRequired")) = "Y" Then
		f_CheckIsThisRequired = true
	Else
		f_CheckIsThisRequired = False
	End if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub


Function ValidateScreenData
	If  document.all.MAILBOX_ID.innerText = "" then
		MsgBox "Mailbox is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If

	If document.all.TxtLOB.value = "" Then
		MsgBox "LOB is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	
	If document.all.TxtSequence.value <> "" Then
		If IsNumeric(document.all.TxtSequence.value) = false then
			MsgBox "Please enter a number in the Sequence field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		End If
	Else
		MsgBox "Sequence is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If

	ValidateScreenData = true
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.MARID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.MARID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If
	
	If document.all.MATID.value = "" Then
		ExeSave = false
		exit function
	End If
		
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.MARID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "MAILBOX_ASSIGNMENT_RULE_ID"& Chr(129) & document.all.MARID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAILBOX_ID"& Chr(129) & document.all.MAILBOX_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAILBOX_ASSIGNMENT_TYPE_ID"& Chr(129) & document.all.MATID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOB.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQUENCE_NUM"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTING_STATE"& Chr(129) & document.all.TxtRoutingState.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "RULE_ID"& Chr(129) & document.all.RULE_ID.innerText & Chr(129) & "1" & Chr(128)

		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()
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
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>" IsThisRequired="<%=isRequired%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Mailbox Assignment Rule Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="MailboxAssignRuleSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="MATID" value="<%=Request.QueryString("MATID")%>" >
<input type="hidden" NAME="MARID" value="<%=Request.QueryString("MARID")%>" >

<%	
If MARID <> "" Then
	If MARID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		SQLST = "SELECT MA.MAILBOX_ASSIGNMENT_RULE_ID,MA.LOB_CD,MA.SEQUENCE_NUM,MA.ROUTING_STATE, " &_
				"R.RULE_TEXT,MA.MAILBOX_ID, MA.RULE_ID FROM " &_
				"MAILBOX_ASSIGNMENT_RULE MA, RULES R, MAILBOX M WHERE MA.RULE_ID = R.RULE_ID(+) AND " &_
				"MA.MAILBOX_ID = M.MAILBOX_ID(+) AND " &_				
				"MA.MAILBOX_ASSIGNMENT_RULE_ID = " & MARID 
		Set RS = Conn.Execute(SQLST)
		RSROUTINGSTATE= ""
		If Not RS.EOF Then
			RSMAILBOX_ID = RS("MAILBOX_ID")
			RSLOB_CD = RS("LOB_CD")
			RSSEQUENCE= RS("SEQUENCE_NUM")
			If Not IsNull(RS("ROUTING_STATE")) Then RSROUTINGSTATE= CStr(RS("ROUTING_STATE"))
			RSRULE_ID= RS("RULE_ID")
			RSRULE_TEXT= ReplaceQuotesInText(RS("RULE_TEXT"))
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=s_DisplayMsg%></SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" >
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
	<td COLSPAN=3>Mailbox Assignment Rule ID:&nbsp<span id="spanMARID"><%=Request.QueryString("MARID")%></span></td>
</tr>
<tr>
	<td width=75>LOB:<br><select ScrnBtn="TRUE" NAME="TxtLOB" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOB_CD,true)%></select></td>
	<td width=75>Sequence:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=11 TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td width=85>Routing State:<br><SELECT ScrnBtn="TRUE" NAME=TxtRoutingState CLASS=LABEL ONCHANGE="VBScript::Control_OnChange"><OPTION VALUE=""></OPTION><OPTION VALUE="* ">*</OPTION><!--#include file="..\lib\states.asp"--></SELECT></td>
</tr>
</table>
<table class="Label" ONDRAGSTART="return false;">
<tr>
	<td width=40>
	<IMG NAME=BtnAttachRule STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
	<IMG NAME=BtnDetachRule STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
	</td>
	<td nowrap WIDTH=300>Rule Text:&nbsp<SPAN ID=RULE_TEXT CLASS=LABEL TITLE="<%=RSRULE_TEXT%>" ><%=TruncateText(RSRULE_TEXT,RuleTextLen)%></SPAN></td>
	<td>Rule ID:&nbsp<span ID=RULE_ID><%=RSRULE_ID%></span></td>
</tr>	
<tr>
	<td width=40>
	<IMG NAME=BtnAttachMailbox STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Mailbox" ONCLICK="VBScript::AttachMailbox MAILBOX_ID">
	<IMG NAME=BtnDetachMailbox STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Mailbox" OnClick="VBScript::DetachMailbox MAILBOX_ID">
	</td>
	<td>Mailbox ID:<span ID=MAILBOX_ID><%=RSMAILBOX_ID%></span></td>
</tr>
</table>



<%If Not IsNull(RSROUTINGSTATE) Then
	If  CStr(RSROUTINGSTATE) <> "" Then	 %>
<SCRIPT LANGUAGE="VBScript">
	SelectOption document.all.TxtRoutingState,"<%=CStr(RSROUTINGSTATE)%>"
</SCRIPT>
<%	End If
End If  %>

</form>
</body>
</html>


