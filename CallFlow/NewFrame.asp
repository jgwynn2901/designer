<!--#include file="..\lib\common.inc"-->
<%
Response.Expires=0

Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function


Function Swap(InData)
If InData = "on" Then
	Swap = "Y"
Else
	Swap = "N"
End If
End Function

Function Swap2(InData)
If InData <> "" Then
	Swap2 = InData
Else
	Swap2 = "null"
End If
End Function

If Request.QueryString("ACTION") = "SAVE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	NextID = NextPkey("FRAME", "FRAME_ID")
	
	SQL = ""
	SQL = SQL & "INSERT INTO FRAME ("
	SQL = SQL & "FRAME_ID, NAME, "
	SQL = SQL & "TITLE, ATTRIBUTE_PREFIX, "
	'SQL = SQL & "ENABLEDRULE_ID, VALIDRULE_ID, "
	SQL = SQL & "MODAL_FLG, " 'ENTRY_ACTION_ID, ACTION_ID, 
	SQL = SQL & "HELPSTRING, "
	SQL = SQL & "DESCRIPTION, TYPE, "
	SQL = SQL & "MAXPAGERESULTROWS, ONEROWAUTOSELECT_FLG ) VALUES ("
	SQL = SQL & NextID & ", " 
	SQL = SQL & "'" & Request.Form("NAME") & "', "
	SQL = SQL & "'" & Request.Form("TITLE") & "', " 
	SQL = SQL & "'" & Request.Form("ATTRIBUTE_PREFIX") & "', " 
	'SQL = SQL & Swap2(Request.Form("ENABLEDRULE_ID")) & ", " 
	'SQL = SQL & Swap2(Request.Form("VALIDRULE_ID")) & ", " 
	SQL = SQL & "'" & Swap(Request.Form("MODAL_FLG")) & "', " 
	'SQL = SQL & Swap2(Request.Form("ENTRY_ACTION_ID")) & ", " 
	'SQL = SQL & Swap2(Request.Form("ACTION_ID")) & ", " 
	SQL = SQL & "'" & Request.Form("HELPSTRING") & "', " 
	SQL = SQL & "'" & Request.Form("DESCRIPTION") & "', " 
	SQL = SQL & "'" & Request.Form("TYPE") & "', "
	
	If IsNumeric(Request.Form("MAXPAGERESULTROWS")) Then
		SQL = SQL & Request.Form("MAXPAGERESULTROWS") & ", " 
	Else
		SQL = SQL & "null, " 
	End If
	
	SQL = SQL & "'" & Swap(Request.Form("ONEROWAUTOSELECT_FLG")) & "') "  
	Set RS = Conn.Execute(SQL)
		
	SQL2 = ""
	SQL2 = SQL2 & "INSERT INTO FRAME_ORDER (FRAME_ID, CALLFLOW_ID" 
	SQL2 = SQL2 & ", SEQUENCE "
	SQL2 = SQL2 & ",TITLE, ATTRIBUTE_PREFIX, ENABLEDRULE_ID, VALIDRULE_ID, "
	SQL2 = SQL2 & "MODAL_FLG, ENTRY_ACTION_ID, ACTION_ID, HELPSTRING, "
	SQL2 = SQL2 & "DESCRIPTION, TYPE, SQLSELECT, SQLFROM, SQLWHERE, SQLORDERBY, "
	SQL2 = SQL2 & "MAXPAGERESULTROWS, ONEROWAUTOSELECT_FLG "
	SQL2 = SQL2 & ") VALUES ("
	SQL2 = SQL2 & NextID & ", " 
	SQL2 = SQL2 & Request.Form("CALLFLOW_ID") 
	SQL2 = SQL2 & ",0"
	SQL2 = SQL2 & ",'-999999999', '-999999999', -999999999, -999999999, "
	SQL2 = SQL2 & "'U', -999999999, -999999999, '-999999999', "
	SQL2 = SQL2 & "'-999999999', '-999999999', '-999999999', '-999999999', '-999999999', '-999999999', "
	SQL2 = SQL2 & "-999999999, 'U' "
	SQL2 = SQL2 & ")" 
	Set RS2 = Conn.Execute(SQL2)
	
End If

%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub BtnSave_onclick
Dim ErrMsg
ErrMsg = "" 
	If document.all.NAME.value = "" Then
		ErrMsg = ErrMsg & "Name cannot be null" & VbCrlf
	End If
	If document.all.TITLE.value = "" Then
		ErrMsg = ErrMsg & "Title cannot be null" & VbCrlf
	End If
	If Not Isnumeric(document.all.MAXPAGERESULTROWS.value) AND document.all.MAXPAGERESULTROWS.value <> "" Then
		ErrMsg = ErrMsg & "Max Result Rows Per Page must be numeric" & VbCrlf
	End If
	'If Not Isnumeric(document.all.SEQUENCE.value) AND document.all.SEQUENCE.value <> "" Then
	'	ErrMsg = ErrMsg & "Sequence must be numeric" & VbCrlf
	'End If
	
	If ErrMsg = "" Then
		FrmFrame.Submit()
	Else
		MsgBox ErrMsg, 0 , "FNSDesigner"
	End If
End Sub

Sub Window_OnLoad
	<% If Request.QueryString("ACTION") = "SAVE" Then %>
		top.frames.location.href = "CallFlow-f.asp?CFID=<%= Request.Form("CALLFLOW_ID") %>&FRAMEID=<%= NextID %>" 
	<% End If %>
	<% If Request.QueryString("NEW") = "TRUE" Then %>
		top.frames.location.href = "NewFrame.asp?CFID=<%= Request.QueryString("CFID") %>"
	<% End If %>

End Sub

Sub BtnCancel_onclick
	top.frames.location.href = "CallFlow-f.asp?CFID=<%= Request.QueryString("CFID") %>"
End Sub

-->
</SCRIPT>
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
	this.SelectButtonLabel = "<U>A</U>ttach";
	this.SelectButtonAccessKey = "A";
}

function CRuleEditorObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
}
var RuleEditorObj = new CRuleEditorObj();
var RuleSearchObj = new CRuleSearchObj();
</script>
</HEAD>
<BODY BGCOLOR=#d6cfbd topmargin=5 rightmargin=0 leftmargin=0 CanDocUnloadNowInf=NO>
<FORM NAME=FrmFrame ACTION="NewFrame.ASP?ACTION=SAVE&CFID=<%= Request.QueryString("CFID") %>" METHOD=POST>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 New Call Flow Frame
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

<TABLE>
<TR>
<TD>
<TABLE>
<TR>
<TD COLSPAN=6 CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT MAXLENGTH=255 NAME="NAME" SIZE=95 OnChange="SetDirty" OnKeyPress="SetDirty" CLASS=LABEL></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL COLSPAN=2>Title:<BR><INPUT TYPE=TEXT SIZE=95 MAXLENGTH=80 NAME="TITLE" OnChange="SetDirty" OnKeyPress="SetDirty" CLASS=LABEL></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL>Help String:<BR><INPUT MAXLENGTH=2000 TYPE=TEXT OnChange="SetDirty" OnKeyPress="SetDirty" SIZE=95 NAME="HELPSTRING" CLASS=LABEL></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT MAXLENGTH=255 OnChange="SetDirty" OnKeyPress="SetDirty" SIZE=95 NAME="DESCRIPTION" CLASS=LABEL ></TD>
</TR>
<TR>
<TD CLASS=LABEL>Attribute Prefix:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=40 OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=40 NAME="ATTRIBUTE_PREFIX"></TD>
<TD CLASS=LABEL>Type:<BR><INPUT TYPE=TEXT SIZE=20 MAXLENGTH=30 OnChange="SetDirty" OnKeyPress="SetDirty" NAME="TYPE" CLASS=LABEL></TD>
<!--<TD CLASS=LABEL>Sequence:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=10 OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=10 NAME="SEQUENCE" ></TD>-->
<TD CLASS=LABEL>Max Result Rows:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=8 OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=10 NAME="MAXPAGERESULTROWS" ></TD>
</TR>
</TABLE>
</TD>
<TD VALIGN=TOP>
<TABLE>
<TR>
<TD>&nbsp;</TD>
</TR>
<TR>
<TD CLASS=LABEL><INPUT TYPE=CHECKBOX OnChange="SetDirty" OnClick="SetDirty" NAME="MODAL_FLG"></TD>
<TD CLASS=LABEL>Modal?</TD>
</TR>
<TR>
<TD CLASS=LABEL><INPUT TYPE=CHECKBOX OnChange="SetDirty" OnClick="SetDirty" NAME="ONEROWAUTOSELECT_FLG"></TD>
<TD CLASS=LABEL>One row auto select?</TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME=CALLFLOW_ID VALUE="<%= Request.QueryString("CFID") %>">
<BR><BR>
&nbsp;<BUTTON NAME=BtnSave CLASS=STDBUTTON ACCESSKEY="S"><U>S</U>ave</BUTTON>&nbsp;
&nbsp;<BUTTON NAME=BtnCancel CLASS=STDBUTTON ACCESSKEY="S"><U>C</U>ancel</BUTTON>&nbsp;
</FORM>
<BR>
<% If Request.querystring("STATUS") = "TRUE" Then %>
<TABLE>
<TR>
<TD CLASS=LABEL><IMG SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER=0 TITLE="Status Report" NOWRAP VALIGN=BOTTOM NAME=BtnStatus ID=BtnStatus></TD>
<TD CLASS=LABEL><FONT COLOR=MAROON>Saved! 
<% If Request.QueryString("WARNINGS") = "NOSEQ" Then %>
&nbsp; Warning: Frame saved with no sequence.
<% End If %>
</TD>
</TR>
</TABLE>
<% End If %>

</BODY>
</HTML>
