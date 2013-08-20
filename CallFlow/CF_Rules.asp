<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<%
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "CF_Rules.asp"
End If

If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"

Function Swap(InData)
If InData <> "" Then
	Swap = InData
Else
	Swap = "null"
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


If Len(Request.QueryString("FRAMEID")) < 1 OR IsNumeric(Request.QueryString("FRAMEID")) = False Then
	Session("ErrorMessage") = "On page " &  Request.ServerVariables("SCRIPT_NAME") & " QueryString FRAMEID was Null or Not Numeric"
	Response.Redirect "..\directerror.asp"
End If
If Request.QueryString("ACTION") <> "SAVE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT FRAME_ORDER.SEQUENCE, FRAME.*, R1.RULE_TEXT As EnabledRuleText, R2.RULE_TEXT As EntryActionRuleText, "
	SQLST = SQLST & "R3.RULE_TEXT As ActionRuleText, R4.RULE_TEXT As ValidRuleText FROM FRAME_ORDER, "
	SQLST = SQLST & "FRAME, RULES R1,	RULES R2,RULES R3,	RULES R4 WHERE 	FRAME.FRAME_ID = " & Request.QueryString("FRAMEID") & " AND "
	SQLST = SQLST & "FRAME.FRAME_ID = FRAME_ORDER.FRAME_ID AND "
	SQLST = SQLST & "FRAME.ENABLEDRULE_ID = R1.RULE_ID (+) AND "
	SQLST = SQLST & "FRAME.ENTRY_ACTION_ID = R2.RULE_ID (+) AND "
	SQLST = SQLST & "FRAME.ACTION_ID = R3.RULE_ID (+) AND "
	SQLST = SQLST & "FRAME.VALIDRULE_ID = R4.RULE_ID (+)"
	Set RS = Conn.Execute(SQLST)
	
If RS.EOF or isnull(RS) Then
	Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
	Response.redirect	 "..\directerror.asp"
End If

Response.Expires = 0
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

<!--#include file="..\lib\Help.asp"-->

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub SetClean
	document.body.SetAttribute "CanDocUnloadNowInf" , "NO"
End Sub

Sub BtnSave_onclick
	setclean()
	FrmFrame.Submit()
End Sub

Sub Window_Onload
ENABLEDRULE_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("EnabledRuleText"), """", "&quot;") %>"
ACTION_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("ActionRuleText"), """", "&quot;") %>"
ENTRY_ACTION_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("EntryActionRuleText"), """", "&quot;") %>"
VALIDRULE_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("ValidRuleText"), """", "&quot;") %>"
End Sub


Function AttachRule (ID, SPANID)
RID = ID.value
MODE = document.body.getAttribute("ScreenMode")
RuleSearchObj.RID = RID
RuleSearchObj.RIDText = SPANID.innerhtml
RuleSearchObj.Selected = false

If RID = "" Then RID = "NEW"

	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID

showModalDialog  strURL  ,RuleSearchObj ,"dialogWidth:450px;dialogHeight:450px;center"
	SetDirty()
If RuleSearchObj.Selected = true Then
	If RuleSearchObj.RID <> ID.value then
		ID.value = RuleSearchObj.RID
	end if
	SPANID.innerhtml = RuleSearchObj.RIDText
ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
	SPANID.innerhtml = RuleSearchObj.RIDText
End If

End Function

Function DetachRule(ID, SPANID)
MODE = document.body.getAttribute("ScreenMode")
If MODE = "RO" Then 
	Exit Function
End If
	SetDirty()
	ID.value = ""
	SPANID.innerhtml = ""
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

function CanDocUnloadNow()
{
	if (false == confirm("Data has changed. Leave page without saving?"))
		return false;
	else
		return true;
}

var RuleSearchObj = new CRuleSearchObj();
</script>
</head>
</HEAD>
<body BGCOLOR="#d6cfbd" topmargin="5" rightmargin="0" leftmargin="0" ScreenMode="<%= MODE %>" CanDocUnloadNowInf=NO>
<form NAME="FrmFrame" ACTION="CF_RULES.ASP?ACTION=SAVE&amp;FRAMEID=<%= Request.QueryString("FRAMEID") %>" METHOD="POST">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><NOBR>&nbsp;» Frame Rules
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<input TYPE="HIDDEN" NAME="NAME" VALUE="<%= RS("NAME") %>">
<input TYPE="HIDDEN" NAME="TYPE" VALUE="<%= RS("TYPE") %>">
<table>
<tr>
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEnabledRule" TITLE="Attach Rule" OnClick="AttachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEnabledRule" TITLE="Detach Rule" OnClick="DetachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Enabled Rule: 
<span ID="ENABLEDRULE_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="ENABLEDRULE_ID" VALUE="<%= RS("ENABLEDRULE_ID") %>" OnChange="SetDirty()">
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachValidRule" TITLE="Attach Rule" OnClick="AttachRule VALIDRULE_ID, VALIDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachValidRule" TITLE="Detach Rule" OnClick="DetachRule VALIDRULE_ID, VALIDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Valid Rule: 
<span ID="VALIDRULE_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="VALIDRULE_ID" VALUE="<%= RS("VALIDRULE_ID") %>" OnChange="SetDirty()">
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEActionRule" TITLE="Attach Rule" OnClick="AttachRule ENTRY_ACTION_ID, ENTRY_ACTION_ID_TEXT " WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEActionRule" TITLE="Detach Rule" OnClick="DetachRule ENTRY_ACTION_ID,ENTRY_ACTION_ID_TEXT " WIDTH="16" HEIGHT="16">
Entry Action Rule:
<span ID="ENTRY_ACTION_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="ENTRY_ACTION_ID" VALUE="<%= RS("ENTRY_ACTION_ID") %>" OnChange="SetDirty()" >
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachActionRule" TITLE="Attach Rule" OnClick="AttachRule ACTION_ID, ACTION_ID_TEXT " WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachActionRule" TITLE="Detach Rule" OnClick="DetachRule ACTION_ID, ACTION_ID_TEXT  " WIDTH="16" HEIGHT="16">
Action Rule:
<span ID="ACTION_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="ACTION_ID" VALUE="<%= RS("ACTION_ID") %>" OnChange="SetDirty()"></td>
</tr>
</table>
<br>
&nbsp;<button NAME="BtnSave" <% If MODE="RO" Then Response.Write(" DISABLED ") %> CLASS="STDBUTTON" ACCESSKEY="S"><u>S</u>ave</button>&nbsp;
<br>
<% If Request.querystring("STATUS") = "TRUE" Then %>
<table>
<tr>
<td CLASS="LABEL"><img SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER="0" TITLE="Status Report" NOWRAP VALIGN="BOTTOM" NAME="BtnStatus" ID="BtnStatus" WIDTH="16" HEIGHT="16"></td>
<td CLASS="LABEL"><font COLOR="MAROON">Saved! </td>
</tr>
</table>
<% End If %>
</form>
</body>
</html>
<% Else
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	
	SQLST = SQLST & "UPDATE FRAME SET NAME='" & Request.Form("NAME") & "', "
	
	If Request.Form("ENABLEDRULE_ID") <> "" Then
		SQLST = SQLST & "ENABLEDRULE_ID =" & Request.Form("ENABLEDRULE_ID") & ", "
	Else
		SQLST = SQLST & "ENABLEDRULE_ID =null, "
	End If
	If Request.Form("ACTION_ID") <> "" Then
		SQLST = SQLST & "ACTION_ID=" & Request.Form("ACTION_ID") & ", "
	Else
		SQLST = SQLST & "ACTION_ID=null, "
	End If
	
	If Request.Form("ENTRY_ACTION_ID") <> "" Then
		SQLST = SQLST & "ENTRY_ACTION_ID=" & Request.Form("ENTRY_ACTION_ID") & ", "
	Else
		SQLST = SQLST & "ENTRY_ACTION_ID=null, "
	End If
	
	If Request.Form("VALIDRULE_ID") <> "" Then
		SQLST = SQLST & "VALIDRULE_ID=" & Request.Form("VALIDRULE_ID") & ", "
	Else
		SQLST = SQLST & "VALIDRULE_ID=null, "
	End If
	
	SQLST = SQLST & "TYPE='" & Request.Form("TYPE") & "' "
	SQLST = SQLST & "WHERE FRAME_ID=" & Request.QueryString("FRAMEID")
	Set RS = Conn.Execute(SQLST)
	
	Response.Redirect "CF_RULES.asp?STATUS=TRUE&FRAMEID=" & Request.QueryString("FRAMEID")
End If
%>

