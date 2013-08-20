<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "CF_General.asp"
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
	SQLST = SQLST & "SELECT FRAME_ORDER.SEQUENCE, FRAME.* FROM  "
	SQLST = SQLST & "FRAME, FRAME_ORDER WHERE 	FRAME.FRAME_ID = " & Request.QueryString("FRAMEID") & " AND "
	SQLST = SQLST & "FRAME.FRAME_ID = FRAME_ORDER.FRAME_ID "
	Set RS = Conn.Execute(SQLST)
	
If RS.EOF or isnull(RS) Then
	Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
	Response.redirect	 "..\directerror.asp"
End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
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
	If document.all.NAME.value = "" OR document.all.TITLE.value = "" Then
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

Sub Window_Onload
document.all.NAME.focus()
<% If RS("MODAL_FLG") = "Y" Then %>
document.all.MODAL_FLG.checked = True
<% End If %>
<% If RS("ONEROWAUTOSELECT_FLG") = "Y" Then %>
document.all.ONEROWAUTOSELECT_FLG.checked = True
<% End If %>


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
</script>
</HEAD>


<BODY BGCOLOR=#d6cfbd topmargin=5 rightmargin=0 leftmargin=0 CanDocUnloadNowInf=NO>
<FORM NAME=FrmFrame ACTION="CF_GENERAL.ASP?ACTION=SAVE&FRAMEID=<%= Request.QueryString("FRAMEID") %>" METHOD=POST>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<INPUT TYPE="HIDDEN" NAME="WARNINGS">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Call Flow Frame
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD></TD>
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
<TD COLSPAN=6 CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT MAXLENGTH=255 STYLE="TEXT-TRANSFORM:UPPERCASE" NAME="NAME" SIZE=95 OnChange="SetDirty" OnKeyPress="SetDirty" CLASS=LABEL VALUE="<%= RS("NAME") %>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL COLSPAN=2>Title:<BR><INPUT TYPE=TEXT SIZE=95 MAXLENGTH=80 NAME="TITLE" OnChange="SetDirty" OnKeyPress="SetDirty" CLASS=LABEL VALUE="<%= RS("TITLE") %>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL>Help String:<BR><INPUT MAXLENGTH=2000 TYPE=TEXT OnChange="SetDirty" OnKeyPress="SetDirty" SIZE=95 NAME="HELPSTRING" CLASS=LABEL  VALUE="<%= RS("HELPSTRING") %>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>></TD>
</TR>
<TR>
<TD COLSPAN=6 CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT MAXLENGTH=255 OnChange="SetDirty" OnKeyPress="SetDirty" SIZE=95 NAME="DESCRIPTION" CLASS=LABEL VALUE="<%= RS("DESCRIPTION") %>" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>></TD>
</TR>
<TR>
<TD CLASS=LABEL>Attribute Prefix:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=40 OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=40 NAME="ATTRIBUTE_PREFIX" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> VALUE="<%= RS("ATTRIBUTE_PREFIX") %>"></TD>
<TD CLASS=LABEL>Type:<BR><INPUT TYPE=TEXT SIZE=20 MAXLENGTH=30 OnChange="SetDirty" OnKeyPress="SetDirty" NAME="TYPE" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS=LABEL VALUE="<%= RS("TYPE") %>"></TD>
<!--<TD CLASS=LABEL>Sequence:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=10 OnChange="SetDirty" OnKeyPress="SetDirty" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> MAXLENGTH=10 NAME="SEQUENCE" VALUE="<%= RS("SEQUENCE") %>"></TD>-->
<TD CLASS=LABEL>Max Result Rows:<BR><INPUT CLASS=LABEL TYPE=TEXT SIZE=13 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> OnChange="SetDirty" OnKeyPress="SetDirty" MAXLENGTH=10 NAME="MAXPAGERESULTROWS" VALUE="<%= RS("MAXPAGERESULTROWS") %>"></TD>
</TR>
</TABLE>
</TD>
<TD VALIGN=TOP>
<TABLE>
<TR>
<TD>&nbsp;</TD>
</TR>
<TR>
<TD CLASS=LABEL><INPUT TYPE=CHECKBOX OnChange="SetDirty" OnClick="SetDirty" NAME="MODAL_FLG" <% If MODE="RO" Then Response.Write(" DISABLED ") %>></TD>
<TD CLASS=LABEL>Modal?</TD>
</TR>
<TR>
<TD CLASS=LABEL><INPUT TYPE=CHECKBOX OnChange="SetDirty" <% If MODE="RO" Then Response.Write(" DISABLED ") %> OnClick="SetDirty" NAME="ONEROWAUTOSELECT_FLG"></TD>
<TD CLASS=LABEL>One row auto select?</TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>

<BR>
&nbsp;<BUTTON NAME=BtnSave <% If MODE="RO" Then Response.Write(" DISABLED ") %> CLASS=STDBUTTON ACCESSKEY="S"><U>S</U>ave</BUTTON>&nbsp;
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
</FORM>

</BODY>
</HTML>
<% Else
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	FR_NAME = UCase(replace(trim(Request.Form("NAME"))," ", ""))
	SQLST = SQLST & "UPDATE FRAME SET NAME='" & FR_NAME & "', "
	SQLST = SQLST & "TITLE='" & Replace(Request.Form("TITLE"),"'", "''") & "', "
	SQLST = SQLST & "ATTRIBUTE_PREFIX='" & Replace(Request.Form("ATTRIBUTE_PREFIX")," ", "") & "', "
	If Request.Form("MODAL_FLG") = "on" Then 
		SQLST = SQLST & "MODAL_FLG='Y', "
	Else
		SQLST = SQLST & "MODAL_FLG='N', "
	End If

	SQLST = SQLST & "HELPSTRING='" & Replace(Request.Form("HELPSTRING"),"'","''") & "', "
	SQLST = SQLST & "DESCRIPTION='" & Replace(Request.Form("DESCRIPTION"),"'","''") & "', "
		
	If Request.Form("ONEROWAUTOSELECT_FLG") = "on" Then 
		SQLST = SQLST & "ONEROWAUTOSELECT_FLG='Y', "
	Else
		SQLST = SQLST & "ONEROWAUTOSELECT_FLG='N', "
	End If
	
	If Request.Form("MAXPAGERESULTROWS") <> "" Then
		SQLST = SQLST & "MAXPAGERESULTROWS=" & Request.Form("MAXPAGERESULTROWS") & ", "
	Else
		SQLST = SQLST & "MAXPAGERESULTROWS =null, "
	End If
	
	SQLST = SQLST & "TYPE='" & Replace(Request.Form("TYPE"),"'","''") & "' "
	SQLST = SQLST & "WHERE FRAME_ID=" & Request.QueryString("FRAMEID")
	Set RS = Conn.Execute(SQLST)
	
	
	'SQL2 = ""
	'SQL2 = SQL2 & "UPDATE FRAME_ORDER SET "
	'SQL2 = SQL2 & "SEQUENCE=" & Swap(Request.Form("SEQUENCE"))
	'SQL2 = SQL2 & " WHERE FRAME_ID="& Request.QueryString("FRAMEID")
	'Set RS2 = Conn.Execute(SQL2)
	'AddErrors = ""
	'If Request.Form("SEQUENCE") = "" Then
	'	AddErrors = AddErrors & "&WARNINGS=NOSEQ"
	'End If
	Response.Redirect "CF_General.asp?STATUS=TRUE&FRAMEID=" & Request.QueryString("FRAMEID") & AddErrors
	
End If
%>

