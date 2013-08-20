<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<% Response.Expires=0
Dim NextRPID
If HasViewPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then  	
	Session("NAME") = ""
	Response.Redirect "Override_Layout_Bottom.asp"
End If
If HasModifyPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then MODE = "RO"

Function Swap(InData)
If InData <> "" Then
	Swap = InData
Else
	Swap = "null"
End If
End Function

Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName & "', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function

Function SwapFlg(InData)
	If InData = "on" Then
		SwapFlg = "Y"
	Else
		SwapFlg = "N"
	End If
End Function

Function ReplaceStr(TextIn, SearchStr , Replacement)
   
	Dim WorkText
    Dim Pointer
    
    WorkText = TextIn
    Pointer = InStr(1, WorkText, SearchStr)
    Do While Pointer > 0
      WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
      Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr)
    Loop
    ReplaceStr = WorkText
End Function

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	
If Request.QueryString("ACTION") = "SAVE" AND Request.QueryString("ROUTING_PLAN_ID") = "NEW" Then
	NextRPID = NextPKey("ROUTING_PLAN", "ROUTING_PLAN_ID")
	SQL = ""
	SQL = SQL & "INSERT INTO ROUTING_PLAN (ROUTING_PLAN_ID, "
	SQL = SQL & "ACCNT_HRCY_STEP_ID, LOB_CD, STATE, DESCRIPTION, "
	SQL = SQL & "DESTINATION_TYPE, ENABLERULE_ID, ENABLED_FLG, INPUT_SYSTEM_NAME) VALUES ("
	SQL = SQL & NextRPID & ", "
	SQL = SQL & Request.Form("ACCNT_HRCY_STEP_ID") & ", "
	SQL = SQL & "'" & Request.Form("LOB_CD") & "', "
	SQL = SQL & "'" & Request.Form("STATE") & "', "
	SQL = SQL & "'" & Request.Form("DESCRIPTION") & "', "
	SQL = SQL & "'" & Request.Form("DESTINATION_TYPE") & "', "
	SQL = SQL & Swap(Request.Form("ENABLERULE_ID")) & ", "
	SQL = SQL & "'" & SwapFlg(Request.Form("ENABLED_FLG")) & "', "
	SQL = SQL & "'" & Request.Form("INPUT_SYSTEM_NAME") & "') "
	Set RSInsert = Conn.Execute(SQL)
	'Response.Redirect "RoutingPlanSummary.asp?routing_plan_id=" & NextRPID & "&STATUS=SAVE&AHSID=" & Request.QueryString("AHSID")
End If	
	
If Request.QueryString("ACTION") = "SAVE" AND Request.QueryString("ROUTING_PLAN_ID") <> "NEW" Then
	SQLSAVE = ""
	SQLSAVE = SQLSAVE & "UPDATE ROUTING_PLAN SET "
	SQLSAVE = SQLSAVE & "ACCNT_HRCY_STEP_ID = " & Request.Form("ACCNT_HRCY_STEP_ID") & ","
	SQLSAVE = SQLSAVE & "LOB_CD = '" & Request.Form("LOB_CD") & "',"
	SQLSAVE = SQLSAVE & "STATE = '" & Request.Form("STATE") & "',"
	SQLSAVE = SQLSAVE & "DESCRIPTION = '" & Replace(Request.Form("DESCRIPTION"),"'","''") & "',"
	SQLSAVE = SQLSAVE & "DESTINATION_TYPE = '" & Request.Form("DESTINATION_TYPE") & "',"
	SQLSAVE = SQLSAVE & "ENABLERULE_ID = " & Swap(Request.Form("ENABLEDRULE_ID")) & ","
	SQLSAVE = SQLSAVE & "ENABLED_FLG = '" & SwapFlg(Request.Form("ENABLED_FLG")) & "',"
	SQLSAVE = SQLSAVE & "INPUT_SYSTEM_NAME = '" & Request.Form("INPUT_SYSTEM_NAME") & "' "
	SQLSAVE = SQLSAVE & "WHERE ROUTING_PLAN_ID = " & Request.Form("ROUTING_PLAN_ID")
	Set RS4 = Conn.Execute(SQLSAVE)
	Response.Redirect "RoutingPlanSummary.asp?routing_plan_id=" & Request.Form("ROUTING_PLAN_ID") & "&STATUS=SAVE&AHSID=" & Request.QueryString("AHSID")

End If 

If Request.QueryString("ROUTING_PLAN_ID") <> "NEW" Then
	SQLST = ""
	SQLST = SQLST & "SELECT ROUTING_PLAN.*, RULE_TEXT FROM  RULES, ROUTING_PLAN WHERE ROUTING_PLAN_ID=" & Request.QueryString("routing_plan_id")
	SQLST = SQLST & " AND ROUTING_PLAN.ENABLERULE_ID = RULES.RULE_ID(+)"
	Set RS = Conn.Execute(SQLST)
	If RS.EOF AND RS.BOF Then
		Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
		Response.redirect	 "..\directerror.asp"
	End If	
	ROUTING_PLAN_ID = RS("ROUTING_PLAN_ID")
	ACCNT_HRCY_STEP_ID = RS("ACCNT_HRCY_STEP_ID")
	LOB_CD = RS("LOB_CD")
	STATE = RS("STATE")
	DESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
	DESTINATION_TYPE = RS("DESTINATION_TYPE")
	ENABLERULE_ID = RS("ENABLERULE_ID")
	ENABLED_FLG = RS("ENABLED_FLG")
	INPUT_SYSTEM_NAME = RS("INPUT_SYSTEM_NAME")
	RULE_TEXT = RS("RULE_TEXT")
End If	
%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Routing Plan</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub BtnSave_onclick
ErrMsg = ""
If document.all.DESCRIPTION.value = "" Then
	ErrMsg = ErrMsg & "Description is a required field." & VbCrlf
End If
If document.all.DESTINATION_TYPE.value = "" Then
	ErrMsg = ErrMsg & "Destination Type is a required field." & VbCrlf
End If
If document.all.INPUT_SYSTEM_NAME.value = "" Then
	ErrMsg = ErrMsg & "Input System Name is a required field." & VbCrlf
End If

If ErrMsg = "" Then
	FrmSave.Submit()
Else
	MsgBox ErrMsg, 0, "FNSNet Designer"
End If
End Sub

Sub Window_Onload
<% If Request.QueryString("ACTION") = "SAVE" AND Request.QueryString("ROUTING_PLAN_ID") = "NEW" Then %>
	Parent.location.href = "RoutingPlanSummary-f.asp?ROUTING_PLAN_ID=<%= NextRPID %>&AHSID=<%=Request.QueryString("AHSID") %>"
<% End If %>

StatusSpan.style.color = "#006699"
<% If Request.QueryString("STATUS") = "SAVE" Then %>
	StatusSpan.innerHTML = "Routing Plan Saved"
<% End If %>
	ENABLEDRULE_ID_TEXT.innerHTML = "<%= Mid(ReplaceStr(RULE_TEXT, """", "&quot;"), 1,50) %>"
	ENABLEDRULE_ID_TEXT.title = "<%= Mid(ReplaceStr(RULE_TEXT, """", "&quot;"),1,50) %>"
	document.all.state.value = "<%= STATE %>"
	<% If ENABLED_FLG = "Y" Then %>
	document.all.ENABLED_FLG.checked = True
	<% End If %>
	document.all.INPUT_SYSTEM_NAME.value = "<%= INPUT_SYSTEM_NAME %>"
	document.all.DESTINATION_TYPE.value = "<%= DESTINATION_TYPE %>"
<% If Request.QueryString("ROUTING_PLAN_ID") = "NEW" Then %>	
	document.all.ACCNT_HRCY_STEP_ID.value = "<%= Request.QueryString("AHSID") %>"
<% End If %>

document.all.LOB_CD.value = "<%= LOB_CD %>"
End Sub

Function AttachRule (ID, SPANID)
RID = ID.value
RuleSearchObj.RID = RID
RuleSearchObj.RIDText = SPANID.innerhtml
RuleSearchObj.Selected = false
MODE = document.body.getAttribute("ScreenMode")

If RID = "" Then RID = "NEW"
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID
showModalDialog  strURL  ,RuleSearchObj ,"dialogWidth:450px;dialogHeight:450px;center"
	
If RuleSearchObj.Selected = true Then
	If RuleSearchObj.RID <> ID.value then
		ID.value = RuleSearchObj.RID
	end if
	SPANID.innerhtml = RuleSearchObj.RIDText
	SPANID.Title = RuleSearchObj.RIDText
ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
	SPANID.innerhtml = RuleSearchObj.RIDText
	SPANID.Title = RuleSearchObj.RIDText
End If
End Function

Function DetachRule(ID, SPANID)
<% If MODE="RO" Then Response.Write(" Exit Function ") %>
	ID.value = ""
	SPANID.innerhtml = ""
	SPANID.Title = ""
End Function

Sub BtnNew_onclick
	Parent.location.href = "RoutingPlanSummary-f.asp?ROUTING_PLAN_ID=NEW"
End Sub

Sub StatusRpt_onclick
	msgbox "No other details reported", 0, "FNSDesigner"
End Sub

Sub BtnGrfxBack_OnClick()
<% If Request.QueryString("AHSID") <> "" Then %>
	parent.frames.location.href = "..\AH\NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID")%>&DROPDOWN=ROUTINGPLAN"
<% Else %>
	Parent.parent.frames.location.href = "RoutingPlanSearchModal.asp?CONTAINERTYPE=FRAMEWORK&<%= Request.QueryString %>"
<% End If %>
End Sub

Function AttachNode(ID)
msgbox "WARNING: Changing the A.H.S. ID will attach the routing plan to a different account.", 0 , "FNSDesigner"
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
-->
</script>
<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
	if (false == confirm("Data has changed. Leave page without saving?"))
		return false;
	else
		return true;
}
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CNodeSearchObj()
{
	this.AHSID = "";
	this.Selected = false;
}

var NodeSearchObj = new CNodeSearchObj();
var RuleSearchObj = new CRuleSearchObj();
</script>
</head>
<body BGCOLOR='<%=BODYBGCOLOR%>'  topmargin="0" rightmargin="2" leftmargin="0" CanDocUnloadNowInf="NO" bottommargin=0 ScreenMode="<%= MODE %>">
<!--#include file="..\lib\NavBack.inc"-->
<form NAME="FrmSave" ACTION="RoutingPlanSummary.asp?ACTION=SAVE&ROUTING_PLAN_ID=<%= Request.QueryString("ROUTING_PLAN_ID") %>&AHSID=<%= Request.QueryString("AHSID") %>" METHOD="POST">
<input TYPE="HIDDEN" NAME="ROUTING_PLAN_ID" VALUE="<%= ROUTING_PLAN_ID %>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<input TYPE="HIDDEN" NAME="WARNINGS">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><NOBR>&nbsp;» 
<% If Request.QueryString("ROUTING_PLAN_ID") = "NEW" Then
	Response.Write(" New ")
End If %>Routing Plan Summary
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'>
</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18" VALIGN="BOTTOM" NOWRAP>
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif"  height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</TD>
<td ALIGN="LEFT" VALIGN="CENTER" VALIGN="MIDDLE" NOWRAP>
:<NOBR><SPAN ID="StatusSpan" CLASS=LABEL  STYLE="COLOR:MAROON">Ready</SPAN>
</td>
</tr>
</table>

<LABEL CLASS=LABEL>Routing Plan ID: <%= Request.QueryString("ROUTING_PLAN_ID") %></LABEL>
<table CELLSPACING=2 CELLPADDING=0 BORDER=0>
<tr>
<td>
<table BORDER="0" CELLSPACING=2 CELLPADDING=0 BORDER=1>
<tr>
<td CLASS="LABEL" COLSPAN="9">Description:<br><input TYPE="TEXT" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS="LABEL" NAME="DESCRIPTION" SIZE="80" VALUE="<%= DESCRIPTION %>"></td>
</tr>
<tr>
<td CLASS="LABEL">Dest. Type:<br>
<SELECT STYLE="WIDTH:150" NAME=DESTINATION_TYPE CLASS=LABEL <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
<%= GetValidValuesHTML("DESTINATION_TYPE", "", true) %>
</SELECT>
</TD>
<td CLASS="LABEL"><NOBR>Input System Name:<br>
<SELECT CLASS="LABEL" NAME="INPUT_SYSTEM_NAME" STYLE="WIDTH:100%" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
<OPTION VALUE="">
<OPTION VALUE="FNS NET">FNS NET
<OPTION VALUE="OPEN BASIC">OPEN BASIC
<OPTION VALUE="FNSINETP1">FNSINETP1
</SELECT>
</td>
<TD CLASS=LABEL ALIGN=LEFT VALIGN=MIDDLE>
A.H. Step ID:<BR><INPUT READONLY TYPE=TEXT CLASS=LABEL NAME=ACCNT_HRCY_STEP_ID STYLE="BACKGROUND-COLOR:SILVER" VALUE="<%= ACCNT_HRCY_STEP_ID %>" SIZE=10>
</TD>
<TD valign=bottom ALIGN=LEFT><IMG src="..\Images\attach.GIF" TITLE="Attach Account Hierarchy Step" STYLE="CURSOR:HAND" align=absbottom OnClick='AttachNode ACCNT_HRCY_STEP_ID'></TD>
</tr>
<tr>
<td CLASS="LABEL">LOB:<br>
<select STYLE="WIDTH:150" NAME="LOB_CD" CLASS="LABEL" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
<%
	SQLLOB = ""
	SQLLOB = SQLLOB & "SELECT * FROM  LOB"
	Set RS2 = Conn.Execute(SQLLOB)
	Do WHile Not RS2.EOF
%>
<option VALUE="<%= RS2("LOB_CD") %>"><%= RS2("LOB_NAME") %>
<%
RS2.MoveNext
Loop
RS2.Close
%>
</select>
</td>
<td CLASS="LABEL" >State:<br>
<select NAME="STATE" CLASS="LABEL" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
<!--#include file="..\lib\states.asp"-->
</select>
</td>
<td CLASS="LABEL" >Enabled:<br>
<input TYPE="CHECKBOX" NAME="ENABLED_FLG" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
</td>
</tr>
</table>
</td>
<td VALIGN="TOP">
<table>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE="RO" Then Response.Write(" DISABLED ") %>ACCESSKEY="S" NAME="BtnSave"><u>S</u>ave</button></td>
</tr>
<!--
<tr>
<td CLASS="LABEL"><button DISABLED CLASS="STDBUTTON" ACCESSKEY="W" NAME="BtnNew">Ne<u>w</u></button></td>
<tr>
-->
</table>
</td>
</tr>
</table>
<table WIDTH="100%">
<tr WIDTH="100%">
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEnabledRule" TITLE="Attach Rule" OnClick="AttachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEnabledRule" TITLE="Detach Rule" OnClick="DetachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Enabled Rule: 
<span ID="ENABLEDRULE_ID_TEXT" CLASS="LABEL" TITLE=""></span><input TYPE="HIDDEN" NAME="ENABLEDRULE_ID" VALUE="<%= ENABLERULE_ID %>">
</td>
</tr>
</table>
</form>
</body>
</html>
