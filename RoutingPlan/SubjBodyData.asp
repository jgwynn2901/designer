<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then  	
	Session("NAME") = ""
	Response.Redirect "Override_Layout_Bottom.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

RuleTextLen = 30
cTRANSMISSION_SEQ_STEP_ID = Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
cOUTPUT_SUBJECT_BODY_ID  = Request.QueryString("OSBID")
If Request.QueryString("STATUS") = "UPDATE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = SQL & "SELECT * FROM OUTPUT_SUBJECT_BODY WHERE OUTPUT_SUBJECT_BODY_ID=" & cOUTPUT_SUBJECT_BODY_ID
	set rs = conn.Execute(SQL)
	cSUBJECT_FILE = RS("SUBJECT_FILE_NAME")
	cBODY_FILE = RS("BODY_FILE_NAME")
	RS.close
	Conn.Close 
	set RS=nothing
	set Conn=nothing
End If	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	document.all.StatusSpan.Style.Color = "#006699"
End Sub

</script>
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" ScreenMode="<%= MODE %>">
<form NAME="FrmSave" TARGET="hiddenPage" ACTION="SubjBodySave.asp?STATUS=<%= Request.QueryString("STATUS") %>" METHOD="POST" ID="Form1">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» <%= Request.QueryString("STATUS") %> Email Subject and Body Template Files
</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0" ID="Table3">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="StatusSpan" CLASS="LABEL" STYLE="COLOR:MAROON">Ready</span>
</td>
</tr>
</table>
<input TYPE="HIDDEN" NAME="TRANSMISSION_SEQ_STEP_ID" VALUE="<%= cTRANSMISSION_SEQ_STEP_ID %>" ID="Hidden1">
<input TYPE="HIDDEN" NAME="OUTPUT_SUBJECT_BODY_ID" VALUE="<%= cOUTPUT_SUBJECT_BODY_ID %>" ID="Hidden2">
<table ID="Table4">
<tr>
<td CLASS="LABEL">Subject File Name:<br>
<input TYPE="TEXT" CLASS="LABEL" NAME="SUBJECT_FILE" SIZE="80" MAXLENGTH="2000" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> VALUE="<%=cSUBJECT_FILE%>" ID="Text1"></td>
</tr>

<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">Body File Name:<br>
<input TYPE="TEXT" SIZE="80" CLASS="LABEL" NAME="BODY_FILE" MAXLENGTH="2000" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> VALUE="<%=cBODY_FILE%>" ID="Text2"></td>
</tr>
<tr></tr>
<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">* Subject and Body File Names can be evaluation expressions that are resolved into the file names<br>Example:<br> $iif("~SOME_ATTRIBUTE~" = "Y", "EmailSubject1.txt", "EmailSubject2.txt")<br>
</tr>
</table>
</form>
</body>
</html>


