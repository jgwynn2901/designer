<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
Response.Expires = 0
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then    
	Session("NAME") = ""
	Response.Redirect "CF_FrameOrder.asp"
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
	SQLST = SQLST & "SELECT FRAME_ORDER.*, R1.RULE_TEXT As EnabledRuleText, R2.RULE_TEXT As EntryActionRuleText, "
	SQLST = SQLST & "R3.RULE_TEXT As ActionRuleText, R4.RULE_TEXT As ValidRuleText FROM RULES R1, RULES R2, RULES R3, RULES R4, "
	SQLST = SQLST & "FRAME_ORDER WHERE FRAME_ID = " & Request.QueryString("FRAMEID") & " AND "
	SQLST = SQLST & "FRAME_ORDER.CALLFLOW_ID = " & Request.QueryString("CALLFLOW_ID") & " AND "
	SQLST = SQLST & "FRAME_ORDER.ENABLEDRULE_ID = R1.RULE_ID (+) AND "
	SQLST = SQLST & "FRAME_ORDER.ENTRY_ACTION_ID = R2.RULE_ID (+) AND "
	SQLST = SQLST & "FRAME_ORDER.ACTION_ID = R3.RULE_ID (+) AND "
	SQLST = SQLST & "FRAME_ORDER.VALIDRULE_ID = R4.RULE_ID (+)"
	Set RS = Conn.Execute(SQLST)

If RS.EOF or isnull(RS) Then
	Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
	Response.redirect	 "..\directerror.asp"
End If
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

<!--#include file="..\lib\Help.asp"-->

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub BtnSave_onclick
Dim ErrMsg
ErrMsg = "" 
	If Not Isnumeric(document.all.MAXPAGERESULTROWS.value) AND document.all.MAXPAGERESULTROWS.value <> "" Then
		ErrMsg = ErrMsg & "Max Result Rows Per Page must be numeric" & VbCrlf
	End If
	If Not Isnumeric(document.all.SEQUENCE.value) AND document.all.SEQUENCE.value <> "" Then
		ErrMsg = ErrMsg & "Sequence must be numeric" & VbCrlf
	End If

	If document.all.ONEROWSPAN.innerhtml = "*" Then
		document.all.ONEROWAUTOSELECT.value = "U"
	else
		If document.all.ONEROWAUTOSELECT_FLG.checked = true Then
			document.all.ONEROWAUTOSELECT.value = "Y"
		Else
			document.all.ONEROWAUTOSELECT.value = "N"
		End If
	End If

	If document.all.MODALSPAN.innerhtml = "*" Then
		document.all.MODAL.value = "U"
	else
		If document.all.MODAL_FLG.checked = true Then
			document.all.MODAL.value = "Y"
		Else
			document.all.MODAL.value = "N"
		End If
	End If


	If ErrMsg = "" Then
		FrmFrame.Submit()
	Else
		MsgBox ErrMsg, 0 , "FNSDesigner"
	End If
End Sub

Sub Window_Onload
document.all.MYTITLE.focus()
<% If RS("MODAL_FLG") = "U" Then %>
document.all.MODALSPAN.innerhtml = "*"
<% Else %>
<% If RS("MODAL_FLG") = "Y" Then %>
document.all.MODAL_FLG.checked = True
<% End If %>
<% End If %>

<% If RS("ONEROWAUTOSELECT_FLG") = "U" Then %>
document.all.ONEROWSPAN.innerhtml = "*"
<% Else %>
<% If RS("ONEROWAUTOSELECT_FLG") = "Y" Then %>
document.all.ONEROWAUTOSELECT_FLG.checked = True
<% End If %>
<% End If %>

<% If RS("ENABLEDRULE_ID") = "-999999999" Then %>
	ENABLEDRULE_ID_TEXT.innerHTML = "*Using frame defined value*"
	ENABLEDRULE_ID_TEXT.style.color = "Maroon"
<% Else %>
	ENABLEDRULE_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("EnabledRuleText"), """", "&quot;") %>"
<% End If %>

<% If RS("ACTION_ID") = "-999999999" Then %>
	ACTION_ID_TEXT.innerHTML = "*Using frame defined value*"
	ACTION_ID_TEXT.style.color = "Maroon"
<% Else %>
	ACTION_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("ActionRuleText"), """", "&quot;") %>"
<% End If %>

<% If RS("ENTRY_ACTION_ID") = "-999999999" Then %>
	ENTRY_ACTION_ID_TEXT.innerHTML = "*Using frame defined value*"
	ENTRY_ACTION_ID_TEXT.style.color = "Maroon"
<% Else %>
	ENTRY_ACTION_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("EntryActionRuleText"), """", "&quot;") %>"
<% End If %>

<% If RS("VALIDRULE_ID") = "-999999999" Then %>
	VALIDRULE_ID_TEXT.innerHTML = "*Using frame defined value*"
	VALIDRULE_ID_TEXT.style.color = "Maroon"
<% Else %>
	VALIDRULE_ID_TEXT.innerHTML = "<%= ReplaceStr(RS("ValidRuleText"), """", "&quot;") %>"
<% End If %>

<% If RS("TITLE") = "-999999999" Then %>
document.all.MYTITLE.style.color = "Maroon"
<% End If %>

<% If RS("SQLSELECT") = "-999999999" Then %>
document.all.SQLSELECT.style.color = "Maroon"
<% End If %>
<% If RS("SQLFROM") = "-999999999" Then %>
document.all.SQLFROM.style.color = "Maroon"
<% End If %>
<% If RS("SQLWHERE") = "-999999999" Then %>
document.all.SQLWHERE.style.color = "Maroon"
<% End If %>
<% If RS("SQLORDERBY") = "-999999999" Then %>
document.all.SQLORDERBY.style.color = "Maroon"
<% End If %>

<% If RS("HELPSTRING") = "-999999999" Then %>
document.all.HELPSTRING.style.color = "Maroon"
<% End If %>

<% If RS("DESCRIPTION") = "-999999999" Then %>
document.all.DESCRIPTION.style.color = "Maroon"
<% End If %>

<% If RS("ATTRIBUTE_PREFIX") = "-999999999" Then %>
document.all.ATTRIBUTE_PREFIX.style.color = "Maroon"
<% End If %>

<% If RS("TYPE") = "-999999999" Then %>
document.all.MYTYPE.style.color = "Maroon"
<% End If %>

<% If RS("MAXPAGERESULTROWS") = "-999999999" Then %>
document.all.MAXPAGERESULTROWS.style.color = "Maroon"
<% End If %>

<% If RS("ONEROWAUTOSELECT_FLG") = "U" Then %>
document.all.ONEROWSPAN.style.color = "Maroon"
document.all.ONEROWSPAN.innerhtml = "*"
<% End If %>

<% If RS("MODAL_FLG") = "U" Then %>
document.all.MODALSPAN.innerhtml = "*"
document.all.MODALSPAN.style.color = "Maroon"
<% End If %>

End Sub

Function AttachRule (ID, SPANID)
RID = ID.value
MODE = document.body.getAttribute("ScreenMode")
RuleSearchObj.RID = RID
RuleSearchObj.RIDText = SPANID.innerhtml
RuleSearchObj.Selected = false

If RID = "" Then RID = "NEW"

	If (RID = "NEW" OR RID="-999999999" ) And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSDesigner"
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
	SetDirty()
	ID.value = ""
	SPANID.innerhtml = ""
End Function

Function SetUnknown(ID, SPANID, ELTYPE, flgnm)
<% If MODE="RO" Then Response.write(" Exit Function ") %>
Select Case ELTYPE
	Case "RULE"
		ID.value = "-999999999"
		SPANID.innerHtml = "*Using frame defined value*"
		SPANID.Title = "*Using frame defined value*"
		SPANID.style.color = "Maroon"
	Case "FLAG"
		Select Case flgnm
			Case "ONEROWAUTOSELECT_FLG"
				ONEROWSPAN.innerHTML = "*"
				ONEROWSPAN.style.color = "Maroon"
			Case "MODAL_FLG"
				MODALSPAN.innerHTML = "*"
				MODALSPAN.style.color = "Maroon"
		End Select
	Case "TEXT"
		ID.Value = "-999999999"
		ID.Style.color = "Maroon"
End Select
End Function

Sub ChngColor(ID)
	ID.Style.color = "Black"
End Sub

Sub MODAL_FLG_onclick
	document.all.MODALSPAN.innerhtml = ""
End Sub

Sub ONEROWAUTOSELECT_FLG_onclick
	document.all.ONEROWSPAN.innerhtml = ""
End Sub
-->
</SCRIPT>
<script LANGUAGE="JavaScript">
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

</HEAD>
<BODY BGCOLOR=#d6cfbd topmargin=5 rightmargin=0 leftmargin=0 CanDocUnloadNowInf=NO ScreenMode="<%= MODE %>">
<FORM NAME=FrmFrame ACTION="CF_FrameOrder.ASP?ACTION=SAVE&FRAMEID=<%= Request.QueryString("FRAMEID") %>&CALLFLOW_ID=<%= Request.QueryString("CALLFLOW_ID") %>" METHOD=POST>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<INPUT TYPE="HIDDEN" NAME="WARNINGS">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Frame Order
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Frame Order.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME=CALLFLOW_ID VALUE="<%=Request.QueryString("CALLFLOW_ID") %>">
<TABLE>
<TR>
<TD>
<TABLE>
<TR>
<TD COLSPAN=6 CLASS=LABEL COLSPAN=2>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN2 OnClick='SetUnknown MYTITLE,MYTITLE, "TEXT", ""'>
Title:<BR><INPUT MAXLENGTH=80 TYPE=TEXT OnChange="SetDirty" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> OnKeyPress="SetDirty" OnKeyDown= 'ChngColor(MYTITLE)'  SIZE=95 NAME="MYTITLE" CLASS=LABEL  VALUE="<%= RS("TITLE") %>" OnKeyDown= 'ChngColor(MYTITLE)'></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN3 OnClick='SetUnknown HELPSTRING,HELPSTRING, "TEXT", ""'>
Help String:<BR><INPUT MAXLENGTH=2000 TYPE=TEXT <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> OnChange="SetDirty" OnKeyPress="SetDirty" OnKeyDown= 'ChngColor(HELPSTRING)'  SIZE=95 NAME="HELPSTRING" CLASS=LABEL  VALUE="<%= RS("HELPSTRING") %>"></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN4 OnClick='SetUnknown DESCRIPTION,DESCRIPTION, "TEXT", ""'>
Description:<BR><INPUT TYPE=TEXT MAXLENGTH=255 OnChange="SetDirty" SIZE=95  OnKeyPress="SetDirty" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> OnKeyDown= 'ChngColor(DESCRIPTION)'  NAME="DESCRIPTION" CLASS=LABEL VALUE="<%= RS("DESCRIPTION") %>"></TD>
</TR>
<TR>
<TD CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN5 OnClick='SetUnknown ATTRIBUTE_PREFIX,ATTRIBUTE_PREFIX, "TEXT", ""'>
Attribute Prefix:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=40 OnChange="SetDirty" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> OnKeyPress="SetDirty" MAXLENGTH=40 OnKeyDown= 'ChngColor(ATTRIBUTE_PREFIX)'  NAME="ATTRIBUTE_PREFIX" VALUE="<%= RS("ATTRIBUTE_PREFIX") %>"></TD>
<TD CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use attribute defined value" ID=UNKNOWN6 OnClick='SetUnknown MYTYPE,MYTYPE, "TEXT", ""'>
Type:<BR><INPUT TYPE=TEXT SIZE=20 MAXLENGTH=30 OnChange="SetDirty" OnKeyPress="SetDirty" NAME="MYTYPE" OnKeyDown= 'ChngColor(MYTYPE)' <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS=LABEL VALUE="<%= RS("TYPE") %>"></TD>
<TD CLASS=LABEL>
Sequence:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=10 OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=10 NAME="SEQUENCE" VALUE="<%= RS("SEQUENCE") %>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>></TD>
<TD CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN7 OnClick='SetUnknown MAXPAGERESULTROWS,MAXPAGERESULTROWS, "TEXT", ""'>
Max Result Rows:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=13 OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=10 OnKeyDown= 'ChngColor(MAXPAGERESULTROWS)' NAME="MAXPAGERESULTROWS" VALUE="<%= RS("MAXPAGERESULTROWS") %>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>></TD>
</TR>
</TABLE>
</TD>
<TD VALIGN=TOP>
<TABLE>
<TR><TD>&nbsp;</TD>
</TR><TR>
<TD CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN8 OnClick='SetUnknown MODAL_FLG,MODAL_FLG, "FLAG", "MODAL_FLG"'>
<INPUT TYPE=CHECKBOX OnChange="SetDirty" NAME="MODAL_FLG" <% If MODE="RO" Then Response.Write(" DISABLED ") %>></TD>
<TD CLASS=LABEL>
<SPAN ID=MODALSPAN STYLE="COLOR:MAROON"></SPAN>
Modal?
<INPUT TYPE=HIDDEN NAME=MODAL>
</TD>
</TR>
<TR>
<TD CLASS=LABEL>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN9 OnClick='SetUnknown ONEROWAUTOSELECT_FLG,ONEROWAUTOSELECT_FLG, "FLAG", "ONEROWAUTOSELECT_FLG"'>
<INPUT TYPE=CHECKBOX <% If MODE="RO" Then Response.Write(" DISABLED ") %> OnChange="SetDirty" NAME="ONEROWAUTOSELECT_FLG"></TD>
<TD CLASS=LABEL>
<SPAN ID=ONEROWSPAN STYLE="COLOR:MAROON"></SPAN>
One row auto select?
<INPUT TYPE=HIDDEN NAME=ONEROWAUTOSELECT>
</TD>
</TR>
<TR><TD><BR><BR></TD></TR>
<TR></TABLE><TABLE>
<TD CLASS=LABEL>
<BUTTON NAME=BtnSave CLASS=STDBUTTON <% If MODE="RO" Then Response.Write(" DISABLED ") %> ACCESSKEY="S"><U>S</U>ave</BUTTON>
</TD>
</TD></TR></TABLE>
</TD></TR></TABLE>
<TABLE><TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN10 OnClick='SetUnknown SQLSelect,SQLSelect, "TEXT", ""'>
SQLSelect:<BR>
<INPUT TYPE=TEXT NAME=SQLSelect VALUE="<%=Trim(RS("SQLSelect"))%>" CLASS=LABEL SIZE=92 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> MAXLENGTH=255 OnChange="SetDirty"   OnKeyPress="SetDirty" OnKeyDown= 'ChngColor(SQLSelect)' >
</TD>
</TR>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN11 OnClick='SetUnknown SQLFrom,SQLFrom, "TEXT", ""'>
SQLFrom:<BR>
<INPUT TYPE=TEXT NAME=SQLFrom VALUE="<%=Trim(RS("SQLFrom"))%>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS=LABEL SIZE=92 MAXLENGTH=255 OnChange="SetDirty"  OnKeyPress="SetDirty" OnKeyDown= 'ChngColor(SQLFrom)' >
</TD>
</TR>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN12 OnClick='SetUnknown SQLWhere,SQLWhere, "TEXT", ""'>
SQLWhere:<BR>
<INPUT TYPE=TEXT NAME=SQLWhere VALUE="<%= Trim(RS("SQLWhere"))%>" CLASS=LABEL SIZE=92 MAXLENGTH=255 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> OnChange="SetDirty"  OnKeyPress="SetDirty" OnKeyDown= 'ChngColor(SQLWhere)' >
</TD>
</TR>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN13 OnClick='SetUnknown SQLOrderBy,SQLOrderBy, "TEXT", ""'>
SQLOrderBy:<BR>
<INPUT TYPE=TEXT NAME=SQLOrderBy VALUE="<%= Trim(RS("SQLOrderBy"))%>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS=LABEL SIZE=92 MAXLENGTH=255 OnChange="SetDirty"  OnKeyPress="SetDirty" OnKeyDown= 'ChngColor(SQLOrderBy)'>
</TD></TR>
</TABLE>
<table>
<tr>
<td CLASS="LABEL" NOWRAP>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN14 OnClick='SetUnknown ENABLEDRULE_ID,ENABLEDRULE_ID_TEXT, "RULE", ""'>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEnabledRule" TITLE="Attach Rule" OnClick="AttachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEnabledRule" TITLE="Detach Rule" OnClick="DetachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Enabled Rule: 
<span ID="ENABLEDRULE_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="ENABLEDRULE_ID" VALUE="<%= RS("ENABLEDRULE_ID") %>" OnChange="SetDirty()">
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN15 OnClick='SetUnknown VALIDRULE_ID,VALIDRULE_ID_TEXT, "RULE", ""'>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachValidRule" TITLE="Attach Rule" OnClick="AttachRule VALIDRULE_ID, VALIDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachValidRule" TITLE="Detach Rule" OnClick="DetachRule VALIDRULE_ID, VALIDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Valid Rule: 
<span ID="VALIDRULE_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="VALIDRULE_ID" VALUE="<%= RS("VALIDRULE_ID") %>" OnChange="SetDirty()">
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN16 OnClick='SetUnknown ENTRY_ACTION_ID,ENTRY_ACTION_ID_TEXT, "RULE", ""'>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEActionRule" TITLE="Attach Rule" OnClick="AttachRule ENTRY_ACTION_ID, ENTRY_ACTION_ID_TEXT " WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEActionRule" TITLE="Detach Rule" OnClick="DetachRule ENTRY_ACTION_ID,ENTRY_ACTION_ID_TEXT " WIDTH="16" HEIGHT="16">
Entry Action Rule:
<span ID="ENTRY_ACTION_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="ENTRY_ACTION_ID" VALUE="<%= RS("ENTRY_ACTION_ID") %>" OnChange="SetDirty()" >
</td>
</tr>
<tr>
<td CLASS="LABEL" NOWRAP>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND" TITLE="Use frame defined value" ID=UNKNOWN17 OnClick='SetUnknown ACTION_ID,ACTION_ID_TEXT, "RULE", ""'>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachActionRule" TITLE="Attach Rule" OnClick="AttachRule ACTION_ID, ACTION_ID_TEXT " WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachActionRule" TITLE="Detach Rule" OnClick="DetachRule ACTION_ID, ACTION_ID_TEXT  " WIDTH="16" HEIGHT="16">
Action Rule:
<span ID="ACTION_ID_TEXT" CLASS="LABEL"></span><input TYPE="HIDDEN" NAME="ACTION_ID" VALUE="<%= RS("ACTION_ID") %>" OnChange="SetDirty()"></td>
</tr>
</table>
<% If Request.querystring("STATUS") = "TRUE" Then %>
<TABLE>
<TR>
<TD CLASS=LABEL><IMG SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER=0 TITLE="Status Report" NOWRAP VALIGN=BOTTOM NAME=BtnStatus ID=BtnStatus></TD>
<TD CLASS=LABEL><FONT COLOR=MAROON>Saved! 
<% If Request.QueryString("WARNINGS") = "NOSEQ" Then %>
&nbsp; Warning: Frame saved with no sequence.
<% End If %>
<% If Session("StatusMsg") <> "" Then %>
SQL Statement may have syntax errors. <BR>
<%= Session("StatusMsg")  %>
<%
Session("StatusMsg") = ""
Session("StatusNum") = ""
 End If %>
</TD>
</TR>
</TABLE>
<% End If %>
</FORM>
</BODY>
</HTML>
<% Else
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLCHK = ""
	If Request.Form("SQLSELECT") <> "-999999999" AND Request.Form("SQLSELECT") <> "" Then
	SQLCHK = SQLCHK & Replace(Request.Form("SQLSELECT"), "'", "''") & " "
	End If
	If Request.Form("SQLFROM") <> "-999999999" AND Request.Form("SQLFROM") <> "" Then
	SQLCHK = SQLCHK & Request.Form("SQLFROM") & " "
	End If
	If Request.Form("SQLWHERE") <> "-999999999" AND Request.Form("SQLWHERE") <> "" Then
	SQLCHK = SQLCHK & Replace(Request.Form("SQLWHERE"), "'", "''") & " "
	End If
	If Request.Form("SQLORDERBY") <> "-999999999" AND Request.Form("SQLORDERBY") <> "" Then
	SQLCHK = SQLCHK & Replace(Request.Form("SQLORDERBY"), "'", "''")
	End If
	If SQLCHK <> "" Then
		QSQL = ""
		QSQL = QSQL & "{call Designer_2.CheckSQLExpression('" & SQLCHK & "','~', '1' ,{resultset 1, StatusMsg, StatusNum})}"
		Set RS2 = Conn.Execute(QSQL)
		If 	RS2("StatusNum") <> "0" Then
			Session("StatusMsg") = RS2("StatusMsg")
			Session("StatusNum") = RS2("StatusNum")
		End If
	End If
		
	SQLST = ""
	SQLST = SQLST & "UPDATE FRAME_ORDER SET "
	SQLST = SQLST & "TITLE='" & Request.Form("MYTITLE") & "', "
	SQLST = SQLST & "ATTRIBUTE_PREFIX='" & Request.Form("ATTRIBUTE_PREFIX") & "', "
	SQLST = SQLST & "MODAL_FLG='" & Request.Form("MODAL") & "', "
	SQLST = SQLST & "HELPSTRING='" & Request.Form("HELPSTRING") & "', "
	SQLST = SQLST & "DESCRIPTION='" & Request.Form("DESCRIPTION") & "', "
	
	SQLST = SQLST & "SQLSELECT='" & Replace(Request.Form("SQLSELECT"), "'", "''") & "', "
	SQLST = SQLST & "SQLFROM='" & Request.Form("SQLFROM") & "', "
	SQLST = SQLST & "SQLWHERE='" & Replace(Request.Form("SQLWHERE"), "'", "''") & "', "
	SQLST = SQLST & "SQLORDERBY='" & Replace(Request.Form("SQLORDERBY"), "'", "''") & "', "

	SQLST = SQLST & "ONEROWAUTOSELECT_FLG='" & Request.form("ONEROWAUTOSELECT") & "', "
		
	If Request.Form("ENTRY_ACTION_ID") <> "" Then
		SQLST = SQLST & "ENTRY_ACTION_ID=" & Request.Form("ENTRY_ACTION_ID") & ","
	Else
		SQLST = SQLST & "ENTRY_ACTION_ID=null" & ","
	End If
	
	If Request.Form("ACTION_ID") <> "" Then
		SQLST = SQLST & "ACTION_ID=" & Request.Form("ACTION_ID") & ","
	Else
		SQLST = SQLST & "ACTION_ID=null" & ","
	End If
	
	If Request.Form("ENABLEDRULE_ID") <> "" Then
		SQLST = SQLST & "ENABLEDRULE_ID=" & Request.Form("ENABLEDRULE_ID") & ","
	Else
		SQLST = SQLST & "ENABLEDRULE_ID=null" & ","
	End If
	
	If Request.Form("VALIDRULE_ID") <> "" Then
		SQLST = SQLST & "VALIDRULE_ID=" & Request.Form("VALIDRULE_ID") & ","
	Else
		SQLST = SQLST & "VALIDRULE_ID=null" & ","
	End If
	
	If Request.Form("MYTYPE") <> "" Then
		SQLST = SQLST & "TYPE='" & Request.Form("MYTYPE") & "',"
	Else
		SQLST = SQLST & "TYPE=null" & ","
	End If
	
	If Request.Form("MAXPAGERESULTROWS") <> "" Then
		SQLST = SQLST & "MAXPAGERESULTROWS=" & Request.Form("MAXPAGERESULTROWS") & ", "
	Else
		SQLST = SQLST & "MAXPAGERESULTROWS =null, "
	End If
	If Request.Form("SEQUENCE") <> "" Then
		SQLST = SQLST & "SEQUENCE=" & Request.Form("SEQUENCE") & " "
	Else
		SQLST = SQLST & "SEQUENCE =null "
	End If
	SQLST = SQLST & "WHERE FRAME_ID=" & Request.QueryString("FRAMEID") & " AND "
	SQLST = SQLST & "CALLFLOW_ID=" & Request.Form("CALLFLOW_ID") 
	Set RS = Conn.Execute(SQLST)
	Response.Redirect "CF_frameorder.asp?STATUS=TRUE&CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") & "&FRAMEID=" & Request.QueryString("FRAMEID") & AddErrors
	
End If
%>

