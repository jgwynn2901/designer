<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	Dim MCBATID, MCBARID

	MCBATID =  CStr(Request.QueryString("MCBATID"))
	MCBARID =  CStr(Request.QueryString("MCBARID"))
	isRequired = "" & Request.QueryString("RequiredMsg")
	
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
<title>Managed Care Branch Assignment Rule Details</title>
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
		
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_MC_BRANCH_ASSIGNMENT&RID=" & RID
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
	
	strURL = "..\Branch\BranchMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_MC_BRANCH_ASSIGNMENT&SELECTONLY=TRUE&BranchTypeFilter=MANAGEDCARE&BID=" & BID
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

Sub UpdateMCBARID(inMCBARID)
	document.all.MCBARID.value = inMCBARID
	document.all.spanMCBARID.innerText = inMCBARID
	if document.all.spanMCBARID.innerText <> "NEW" then
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

Function GetMCBARID
	if document.all.MCBARID.value <> "NEW" then
		GetMCBARID = document.all.MCBARID.value
	else
		GetMCBARID = ""
	end if 
End Function

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Function f_CheckRequire
	if document.body.getAttribute("IsThisRequired") = "Y" then 
		f_CheckRequire = true
	else
		f_CheckRequire = false
	end if
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

	If  document.all.TxtManagedCareType.value = "" then
		MsgBox "Type is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If

	If document.all.TxtRoutingState.value = "" then
		MsgBox "Routing State is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	If document.all.TxtLOB.value = "" then
		MsgBox "Line of Business is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	ValidateScreenData = true
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.MCBARID.value = "" then
		ExeCopy = false
		exit function
	end if


	document.all.MCBARID.value = "NEW"
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

	If document.all.MCBATID.value = "" Then
		Msgbox "MC BranchAssignmentType ID un-defined."
		ExeSave = false
		exit function
	End If
		
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.MCBARID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "MC_BRANCHASSIGNMENTRULE_ID"& Chr(129) & document.all.MCBARID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BRANCH_ID"& Chr(129) & document.all.BRANCH_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MC_BRANCHASSIGNMENTTYPE_ID"& Chr(129) & document.all.MCBATID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOB.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQUENCE"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTING_STATE"& Chr(129) & document.all.TxtRoutingState.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTING_FIPS"& Chr(129) & document.all.TxtRoutingFIPS.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MANAGED_CARE_TYPE"& Chr(129) & document.all.TxtManagedCareType.value & Chr(129) & "1" & Chr(128)
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
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" IsThisRequired="<%=isRequired%>" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Managed Care Branch Assignment Rule Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="MCBranchAssignRuleSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="MCBATID" value="<%=Request.QueryString("MCBATID")%>" >
<input type="hidden" NAME="MCBARID" value="<%=Request.QueryString("MCBARID")%>" >

<%	
RSROUTING_FIPS = "*"


If MCBARID <> "" Then
	If MCBARID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		SQLST = "SELECT MCBA.MC_BRANCHASSIGNMENTRULE_ID,MCBA.LOB_CD,MCBA.SEQUENCE,MCBA.ROUTING_STATE,MCBA.ROUTING_FIPS,MCBA.MANAGED_CARE_TYPE, " &_
				"R.RULE_TEXT,B.OFFICE_NAME,MCBA.BRANCH_ID, MCBA.RULE_ID FROM " &_
				"MC_BRANCHASSIGNMENTRULE MCBA, RULES R, BRANCH B WHERE MCBA.RULE_ID = R.RULE_ID(+) AND " &_
				"MCBA.BRANCH_ID = B.BRANCH_ID(+) AND " &_				
				"MCBA.MC_BRANCHASSIGNMENTRULE_ID = " & MCBARID 
		Set RS = Conn.Execute(SQLST)
	
		RSROUTING_STATE= ""
		
		If Not RS.EOF Then
			RSBRANCH_ID = RS("BRANCH_ID")
			RSBRANCH_OFFICE_NAME = RS("OFFICE_NAME")
			RSLOB_CD = RS("LOB_CD")
			RSSEQUENCE= RS("SEQUENCE")
			If Not IsNull(RS("ROUTING_STATE")) Then RSROUTING_STATE= CStr(RS("ROUTING_STATE"))
			RSROUTING_FIPS= ReplaceQuotesInText(RS("ROUTING_FIPS"))
			RSMANAGED_CARE_TYPE= ReplaceQuotesInText(RS("MANAGED_CARE_TYPE"))
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
	<td COLSPAN=4>Managed Care Branch Assignment Rule ID:&nbsp<span id="spanMCBARID"><%=Request.QueryString("MCBARID")%></span></td>
</tr>
<tr>
	<td width=75>LOB:<br><select ScrnBtn="TRUE" NAME="TxtLOB" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOB_CD,true)%></select></td>
	<td width=75>Sequence:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=11 TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td width=85>Routing State:<br><SELECT ScrnBtn="TRUE" NAME=TxtRoutingState CLASS=LABEL ONCHANGE="VBScript::Control_OnChange"><OPTION VALUE=""></OPTION><OPTION VALUE="* ">*</OPTION><!--#include file="..\lib\states.asp"--></SELECT></td>
	<td width=85>Routing FIPS:<br><input ScrnInput="TRUE" CLASS="LABEL" SIZE=6 MAXLENGTH=5 TYPE="TEXT" NAME="TxtRoutingFIPS" VALUE="<%=RSROUTING_FIPS%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td width=85>Type:<br><select ScrnBtn="TRUE" NAME="TxtManagedCareType" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><OPTION VALUE="* ">*</OPTION><OPTION VALUE="CERTIFIED">CERTIFIED</OPTION><OPTION VALUE="NOTCERTIFIED">NOTCERTIFIED</OPTION></select></td>
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
	<td>Branch ID:<span ID=BRANCH_ID Name=BRANCH_ID><%=RSBRANCH_ID%></span></td>
</tr>
</table>



<%If Not IsNull(RSROUTING_STATE) Then
	If  CStr(RSROUTING_STATE) <> "" Then	 %>
<SCRIPT LANGUAGE="VBScript">
	SelectOption document.all.TxtRoutingState,"<%=CStr(RSROUTING_STATE)%>"
</SCRIPT>
<%	End If
End If  %>

<%If Not IsNull(RSMANAGED_CARE_TYPE) Then
	If  CStr(RSMANAGED_CARE_TYPE) <> "" Then	 %>
<SCRIPT LANGUAGE="VBScript">
	SelectOption document.all.TxtManagedCareType,"<%=CStr(RSMANAGED_CARE_TYPE)%>"
</SCRIPT>
<%	End If
End If  %>

</form>
</body>
</html>


