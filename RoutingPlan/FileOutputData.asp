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
cOUTPUTDEF_ID = Request.QueryString("ODID")
cOUTPUT_FILE_ID  = Request.QueryString("OFID")
If Request.QueryString("STATUS") = "UPDATE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = SQL & "SELECT * FROM OUTPUT_FILE, RULES WHERE OUTPUT_FILE_ID=" & cOUTPUT_FILE_ID & _
			" AND ENABLE_RULE_ID = RULE_ID(+)"
	set rs = conn.Execute(SQL)
	cCR_FILENAME = RS("REPORT_FILE_NAME")
	cOUTPUT_FILENAME = RS("OUTPUT_FILE_NAME")
	cOUTPUT_FILE_FORMAT = RS("OUTPUT_FILE_FORMAT")
	nENABLE_RULE_ID = RS("ENABLE_RULE_ID")
	RSRULE_TEXT = ReplaceQuotesInText(RS("RULE_TEXT"))
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
	SelectOption document.all.OUTPUT_FORMAT, "<%=cOUTPUT_FILE_FORMAT%>"
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
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_OUTPUT_DEFINITION&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = RuleSearchObj.RID
			document.all.ENABLING_RULE_ID.value = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText(SPANID,inText)
	If Len(inText) < <%=RuleTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid (inText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

</script>
<script LANGUAGE="JScript">
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}

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

var RuleSearchObj = new CRuleSearchObj();
</script>
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" ScreenMode="<%= MODE %>">
<form NAME="FrmSave" TARGET="hiddenPage" ACTION="FileOutputSave.asp?STATUS=<%= Request.QueryString("STATUS") %>" METHOD="POST">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» <%= Request.QueryString("STATUS") %> File Output
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
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="StatusSpan" CLASS="LABEL" STYLE="COLOR:MAROON">Ready</span>
</td>
</tr>
</table>
<input TYPE="HIDDEN" NAME="OUTPUTDEF_ID" VALUE="<%= cOUTPUTDEF_ID %>">
<input TYPE="HIDDEN" NAME="OUTPUT_FILE_ID" VALUE="<%= cOUTPUT_FILE_ID %>" ID="Hidden1">
<input TYPE="HIDDEN" NAME="ENABLING_RULE_ID" VALUE="" ID="Hidden2">
<table>
<tr>
<td CLASS="LABEL">Crystal Reports Filename:<br>
<input TYPE="TEXT" CLASS="LABEL" NAME="CR_FILE" SIZE="35" MAXLENGTH="80" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> VALUE="<%=cCR_FILENAME%>"></td>
</tr>
<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">Output Filename:<br>
<input TYPE="TEXT" SIZE="35" CLASS="LABEL" NAME="OUTPUT_FILE" MAXLENGTH="80" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> VALUE="<%=cOUTPUT_FILENAME%>"></td>
</tr>
<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">Output Format: 
<select name="OUTPUT_FORMAT" CLASS="LABEL" ID="Select1">
	<option VALUE="DOC">Word</option>
	<option VALUE="PDF">PDF</option>
	<option VALUE="HTM">HTML</option>
	<option VALUE="RTF">RTF</option>
	<option VALUE="XLS">Excel</option>
</select>
</tr>
</table>
<table class="Label" ID="Table1" CELLSPACING="3" CELLPADDING="3">
<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">Enabling Rule:<br></td>
</tr>
<tr>
<td>
<img NAME="BtnAttachRule" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
<img NAME="BtnDetachRule" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
</td>
<td width="305" nowrap>Rule Text:&nbsp;<span ID="RULE_TEXT" CLASS="LABEL" TITLE="<%=RSRULE_TEXT%>"><%=TruncateText(RSRULE_TEXT,30)%></span></td>
<td>Rule ID:&nbsp;<span ID="RULE_ID" CLASS="LABEL"><%=nENABLE_RULE_ID%></span></td>
</tr>
</table>
</form>
</body>
</html>


