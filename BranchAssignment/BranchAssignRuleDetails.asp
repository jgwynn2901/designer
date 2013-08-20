<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim BATID, BARID, isRequired
	
	BATID =  CStr(Request.QueryString("BATID"))
	BARID =  CStr(Request.QueryString("BARID"))
	isRequired = Request.QueryString("RequiredMsg")
	
	IF isRequired = "Y" Then
		s_DisplayMsg = "At least 1 BranchAssignmentRule Is Required"
	Else
		s_DisplayMsg = "Ready"
	End If
	
	BranchTextLen = 30
	RuleTextLen = 30
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Branch Assignment Rule Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
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
var BranchSearchObj = new CBranchSearchObj();
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
		
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_BRANCH_ASSIGNMENT&RID=" & RID
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
	
	strURL = "..\Branch\BranchMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_BRANCH_ASSIGNMENT&SELECTONLY=TRUE&BranchTypeFilter=CLAIMHANDLING&BID=" & BID
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

Sub UpdateBARID(inBARID)
	document.all.BARID.value = inBARID
	document.all.spanBARID.innerText = inBARID
	if document.all.spanBARID.innerText <> "NEW" then
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

Function GetBARID
	if document.all.BARID.value <> "NEW" then
		GetBARID = document.all.BARID.value
	else
		GetBARID = ""
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
	If  document.all.BRANCH_ID.innerText = "" then
		MsgBox "Branch is a required field.",0,"FNSNetDesigner"
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
	
	if document.all.BARID.value = "" then
		ExeCopy = false
		exit function
	end if


	document.all.BARID.value = "NEW"
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
	
	If document.all.BATID.value = "" Then
		ExeSave = false
		exit function
	End If
		
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.BARID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "BRANCHASSIGNMENTRULE_ID"& Chr(129) & document.all.BARID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BRANCH_ID"& Chr(129) & document.all.BRANCH_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BRANCHASSIGNMENTTYPE_ID"& Chr(129) & document.all.BATID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOB.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQUENCE"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTINGSTATE"& Chr(129) & document.all.TxtRoutingState.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTINGFIPS"& Chr(129) & document.all.TxtRoutingFIPS.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTINGZIP"& Chr(129) & document.all.TxtRoutingZip.value & Chr(129) & "1" & Chr(128)
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
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Branch Assignment Rule Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="BranchAssignRuleSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="BATID" value="<%=Request.QueryString("BATID")%>" >
<input type="hidden" NAME="BARID" value="<%=Request.QueryString("BARID")%>" >

<%	
RSROUTINGFIPS = "*"
RSROUTINGZIP = "*"


If BARID <> "" Then
	If BARID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		SQLST = "SELECT BA.BRANCHASSIGNMENTRULE_ID,BA.LOB_CD,BA.SEQUENCE,BA.ROUTINGSTATE,BA.ROUTINGFIPS,BA.ROUTINGZIP, " &_
				"R.RULE_TEXT,B.OFFICE_NAME,BA.BRANCH_ID, BA.RULE_ID FROM " &_
				"BRANCHASSIGNMENTRULE BA, RULES R, BRANCH B WHERE BA.RULE_ID = R.RULE_ID(+) AND " &_
				"BA.BRANCH_ID = B.BRANCH_ID(+) AND " &_				
				"BA.BRANCHASSIGNMENTRULE_ID = " & BARID 
		Set RS = Conn.Execute(SQLST)
	
		RSROUTINGSTATE= ""
		
		If Not RS.EOF Then
			RSBRANCH_ID = RS("BRANCH_ID")
			RSBRANCH_OFFICE_NAME = RS("OFFICE_NAME")
			RSLOB_CD = RS("LOB_CD")
			RSSEQUENCE= RS("SEQUENCE")
			If Not IsNull(RS("ROUTINGSTATE")) Then RSROUTINGSTATE= CStr(RS("ROUTINGSTATE"))
			RSROUTINGFIPS= ReplaceQuotesInText(RS("ROUTINGFIPS"))
			RSROUTINGZIP= ReplaceQuotesInText(RS("ROUTINGZIP"))
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
	<td COLSPAN=3>Branch Assignment Rule ID:&nbsp<span id="spanBARID"><%=Request.QueryString("BARID")%></span></td>
</tr>
<tr>
	<td width=75>LOB:<br><select ScrnBtn="TRUE" NAME="TxtLOB" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOB_CD,true)%></select></td>
	<td width=75>Sequence:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=11 TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td width=85>Routing State:<br><SELECT ScrnBtn="TRUE" NAME=TxtRoutingState CLASS=LABEL ONCHANGE="VBScript::Control_OnChange"><OPTION VALUE=""></OPTION><OPTION VALUE="* ">*</OPTION><!--#include file="..\lib\states.asp"--></SELECT></td>
	<td width=85>Routing FIPS:<br><input ScrnInput="TRUE" CLASS="LABEL" SIZE=6 MAXLENGTH=5 TYPE="TEXT" NAME="TxtRoutingFIPS" VALUE="<%=RSROUTINGFIPS%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td width=85>Routing Zip:<br><input ScrnInput="TRUE" CLASS="LABEL" SIZE=10 MAXLENGTH=9 TYPE="TEXT" NAME="TxtRoutingZip" VALUE="<%=RSROUTINGZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
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
	<IMG NAME=BtnAttachBranch STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Branch" ONCLICK="VBScript::AttachBranch BRANCH_ID, BRANCH_DESC,''">
	<IMG NAME=BtnDetachBranch STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Branch" OnClick="VBScript::Detach BRANCH_ID, BRANCH_DESC">
	</td>
	<td nowrap  WIDTH=300>Branch:&nbsp<SPAN ID=BRANCH_DESC CLASS=LABEL TITLE="<%=RSBRANCH_OFFICE_NAME%>" ><%=TruncateText(RSBRANCH_OFFICE_NAME,BranchTextLen)%></SPAN></td>
	<td>Branch ID:<span ID=BRANCH_ID><%=RSBRANCH_ID%></span></td>
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


