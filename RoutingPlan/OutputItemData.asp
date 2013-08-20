<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then  	
	Session("NAME") = ""
	Response.Redirect "Override_Layout_Bottom.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

TRANSMISSION_SEQ_STEP_ID = Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
If Request.QueryString("STATUS") = "UPDATE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = SQL & "SELECT * FROM OUTPUT_ITEM OI, RULES R WHERE OI.RULE_ID = R.RULE_ID(+) AND OUTPUT_ITEM_ID=" & Request.QueryString("OUTPUT_ITEM_ID")
	set rs = conn.Execute(SQL)
	OUTPUT_ITEM_ID = RS("OUTPUT_ITEM_ID")
	TRANSMISSION_SEQ_STEP_ID = RS("TRANSMISSION_SEQ_STEP_ID")
	OUTPUTDEF_ID = RS("OUTPUTDEF_ID")
	SEQUENCE = RS("SEQUENCE")
	ENABLERULE_ID = RS("RULE_ID")
	ENABLERULE_TEXT = RS("RULE_TEXT")
End If	

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

%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</HEAD>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

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


Sub BtnFindOD_onclick
<% If MODE="RO" Then Response.write(" Exit Sub ") %>
	lret = window.showModalDialog( "../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=MODAL"  ,DefinitionObj ,"dialogWidth:450px;dialogHeight:450px;center")
	if DefinitionObj.ODID <> "" Then
		document.all.OUTPUTDEF_ID.value = DefinitionObj.ODID
	end if
End Sub

Sub window_onload
	document.all.StatusSpan.Style.Color = "#006699"
	ENABLERULE_ID_TEXT.innerHTML = "<%= Mid(ReplaceStr(ENABLERULE_TEXT, """", "&quot;"), 1,50) %>"
	ENABLERULE_ID_TEXT.title = "<%= Mid(ReplaceStr(ENABLERULE_TEXT, """", "&quot;"),1,50) %>"
End Sub

-->
</SCRIPT>
<SCRIPT LANGUAGE=JavaScript>
<!--
function COutputDefinitionSearchObj()
{
	this.ODID = "";
	this.ODIDName = "";
	this.Saved = false;	
	this.Selected = false;	
}

function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}

var RuleSearchObj = new CRuleSearchObj();
var DefinitionObj = new COutputDefinitionSearchObj();
//-->
</SCRIPT>
</head>
<body BGCOLOR='<%=BODYBGCOLOR%>'   leftmargin=0 topmargin=0 rightmargin=0 bottommargin=0 ScreenMode="<%= MODE %>">
<FORM NAME="FrmSave" TARGET="hiddenPage" ACTION="SaveOutputItem.asp?STATUS=<%= Request.QueryString("STATUS") %>&ITEM_STEP=<%= Request.QueryString("ITEM_STEP") %>" METHOD=POST>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 <%= Request.QueryString("STATUS") %> Output Item
</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="StatusSpan" CLASS=LABEL  STYLE="COLOR:MAROON">Ready</SPAN>
</td>
</tr>
</table>
<INPUT TYPE=HIDDEN NAME=OUTPUT_ITEM_ID VALUE="<%= OUTPUT_ITEM_ID %>">
<INPUT TYPE=HIDDEN NAME=TRANSMISSION_SEQ_STEP_ID VALUE="<%= TRANSMISSION_SEQ_STEP_ID %>">
<TABLE>
<TR>
<TD CLASS=LABEL>Sequence:<BR><INPUT TYPE=TEXT CLASS=LABEL <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> NAME=SEQUENCE SIZE=10 MAXLENGTH=10 VALUE="<%= SEQUENCE %>"></TD>
</TR>
<TR>
<TD CLASS=LABEL VALIGN=BOTTOM>Output Definition:<BR><INPUT READONLY TYPE=TEXT SIZE=10 CLASS=LABEL NAME=OUTPUTDEF_ID VALUE="<%= OUTPUTDEF_ID %>" STYLE="BACKGROUND-COLOR:SILVER"VALUE="">
<IMG SRC="../IMAGES/Attach.gif" ID=BtnFindOD TITLE="Attach Output Definition" STYLE="CURSOR:HAND" align=absbottom></TD>
</TR>
</TABLE>

<table WIDTH="100%">
<tr WIDTH="100%">
<tr></tr>
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEnabledRule" TITLE="Attach Rule" OnClick="AttachRule ENABLERULE_ID, ENABLERULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEnabledRule" TITLE="Detach Rule" OnClick="DetachRule ENABLERULE_ID, ENABLERULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Enabled/Disabled Rule: 
<span ID="ENABLERULE_ID_TEXT" CLASS="LABEL" TITLE=""></span><input TYPE="HIDDEN" NAME="ENABLERULE_ID" VALUE="<%= ENABLERULE_ID %>">
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>* if no rule is attached, the output item will always be processed</td>
</tr>
</table>

</FORM>
</BODY>
</HTML>


