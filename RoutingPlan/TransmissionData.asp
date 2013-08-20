<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<% Response.Expires=0 
RuleTextLen = 30
If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "Override_Layout_Bottom.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"
StatusLabel = "New"
If Request.QueryString("STATUS") = "UPDATE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	StatusLabel = "Update"
	SQL1 = ""
	SQL1 = SQL1 & "SELECT * FROM TRANSMISSION_SEQ_STEP WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	Set RS2 = Conn.execute(SQL1)
	TRANSMISSION_SEQ_STEP_ID = RS2("TRANSMISSION_SEQ_STEP_ID")
	ROUTING_PLAN_ID = RS2("ROUTING_PLAN_ID")
	SEQUENCE =  RS2("SEQUENCE")
	RETRY_COUNT = RS2("RETRY_COUNT")
	RETRY_WAIT_TIME =  RS2("RETRY_WAIT_TIME")
	DESTINATION_STRING =  RS2("DESTINATION_STRING")
	ALT_DESTINATION_STRING =  RS2("ALT_DESTINATION_STRING")
	TRANSMISSION_TYPE_ID =  RS2("TRANSMISSION_TYPE_ID")
	BATCH_HOLD =  RS2("BATCH_HOLD")
	RS2.Close()
	
	SQL2 = ""
	SQL2 = SQL2 & "SELECT * FROM OUTPUT_FILENAME OTF, RULES R WHERE OTF.RULE_ID = R.RULE_ID AND OTF.TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	Set RS3 = Conn.execute(SQL2)
	if not RS3.EOF then
		OUTPUT_FILENAME_ID = RS3("OUTPUT_FILENAME_ID")
		RULE_ID = RS3("RULE_ID")
		RSRULE_TEXT = RS3("RULE_TEXT")
		DESCRIPTION = RS3("DESCRIPTION")
	end if
	RS3.Close()
	
	if (TRANSMISSION_TYPE_ID = "8" or TRANSMISSION_TYPE_ID = "11" or TRANSMISSION_TYPE_ID = "13") then
		FILENAME_FLG = "Y"
	else
		FILENAME_FLG = "N"
	end if
	
	SQL3 = ""
	SQL3 = SQL3 & "SELECT * FROM OUTPUT_XMLTEMPLATE OTX WHERE OTX.TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	Set RS4 = Conn.execute(SQL3)
	if not RS4.EOF then
		XMLFILE = RS4("FILE_NAME")
		XMLDESCRIPTION = RS4("DESCRIPTION")
	end if
	RS4.Close()
	
	if (TRANSMISSION_TYPE_ID = "3" or TRANSMISSION_TYPE_ID = "14" or TRANSMISSION_TYPE_ID = "13") then
		XML_FILENAME_FLG = "Y"
	else
		XML_FILENAME_FLG = "N"
	end if

End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
<% If Request.QueryString("STATUS") = "UPDATE" Then 
If BATCH_HOLD = "Y" Then
%>
document.all.BATCH_HOLD.checked = true
<% End If %>
document.all.TRANSMISSION_TYPE_ID.value = "<%= TRANSMISSION_TYPE_ID %>"
if document.all.TRANSMISSION_TYPE_ID.value = 8 or _
	document.all.TRANSMISSION_TYPE_ID.value = 11  or _
	document.all.TRANSMISSION_TYPE_ID.value = 13 then
	document.all.RuleDiv.style.visibility="visible"
	document.all.FILENAME_FLG.value = "Y"
else
	document.all.RuleDiv.style.visibility="hidden"
	document.all.FILENAME_FLG.value = "N"
end if

if document.all.TRANSMISSION_TYPE_ID.value = 3 or _
	document.all.TRANSMISSION_TYPE_ID.value = 14 or _
	document.all.TRANSMISSION_TYPE_ID.value = 13 then
	document.all.XMLTemplate.style.visibility="visible"
	document.all.XML_FILENAME_FLG.value = "Y"
else
	document.all.XMLTemplate.style.visibility="hidden"
	document.all.XML_FILENAME_FLG.value = "N"
end if

<% End If %>

document.all.StatusSpan.style.color = "#006699"
End Sub

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "NO"
End Sub

Function checkFileName ()
	if document.all.TRANSMISSION_TYPE_ID.value = 8 or _
		document.all.TRANSMISSION_TYPE_ID.value = 11  or _
		document.all.TRANSMISSION_TYPE_ID.value = 13 then
		document.all.RuleDiv.style.visibility="visible"
		document.all.FILENAME_FLG.value = "Y"
	else
		document.all.RuleDiv.style.visibility="hidden"
		document.all.FILENAME_FLG.value = "N"
	end if
	
	if document.all.TRANSMISSION_TYPE_ID.value = 3 or _
		document.all.TRANSMISSION_TYPE_ID.value = 14 or _
		document.all.TRANSMISSION_TYPE_ID.value = 13 then
		document.all.XMLTemplate.style.visibility="visible"
		document.all.XML_FILENAME_FLG.value = "Y"
	else
		document.all.XMLTemplate.style.visibility="hidden"
		document.all.XML_FILENAME_FLG.value = "N"
	end if
	
	setDirty()
End function


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
			document.all.FILENAME_RULE_ID.value = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	End If

End Function

Function Detach(ID, SPANID)
	if ID.innerText <> "" or SPANID.innerText <> "" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
		document.all.DESCRIPTION.value = ""
		document.all.FILENAME_RULE_ID.value = ""
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
-->
</SCRIPT>
<script LANGUAGE="JScript">
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}

var RuleSearchObj = new CRuleSearchObj();
</script>

</HEAD>
<BODY BGCOLOR=#d6cfbd topmargin=5 rightmargin=0 leftmargin=0 CanDocUnloadNowInf=YES>
<FORM NAME="FrmSave" TARGET="hiddenPage" ACTION="SaveTransmission.asp?STATUS=<%= Request.QueryString("STATUS") %>" METHOD=POST>

<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<INPUT TYPE="HIDDEN" NAME="WARNINGS">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 <%= StatusLabel %> Transmission Sequence Step
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
<INPUT TYPE=HIDDEN NAME="TRANSMISSION_SEQ_STEP_ID" VALUE="<%= TRANSMISSION_SEQ_STEP_ID %>">
<INPUT TYPE=HIDDEN NAME="ROUTING_PLAN_ID" VALUE="<%= Request.QueryString("ROUTING_PLAN_ID") %>">
<INPUT TYPE=HIDDEN NAME="OUTPUT_FILENAME_ID" VALUE="<%= OUTPUT_FILENAME_ID %>">
<INPUT TYPE="HIDDEN" NAME="FILENAME_RULE_ID" VALUE="<%= RULE_ID %>">
<INPUT TYPE="HIDDEN" NAME="FILENAME_FLG" VALUE="<%= FILENAME_FLG %>">
<INPUT TYPE="HIDDEN" NAME="XML_FILENAME_FLG" VALUE="<%= XML_FILENAME_FLG %>">


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


<TABLE>
<TR>
<TD CLASS=LABEL COLSPAN=3>Destination String:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=DESTINATION_STRING SIZE=60 MAXLENGTH=255 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> OnChange="setDirty()" OnKeyPress="setDirty()" VALUE="<%= DESTINATION_STRING %>"></TD>
<TD CLASS=LABEL VALIGN=BOTTOM><INPUT TYPE=CHECKBOX NAME=BATCH_HOLD CLASS=LABEL OnChange="setDirty()" OnKeyPress="setDirty()" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>Batch Hold?</TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Alternate Destination String:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=ALT_DESTINATION_STRING SIZE=60 MAXLENGTH=255 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> OnChange="setDirty()" OnKeyPress="setDirty()" VALUE="<%= ALT_DESTINATION_STRING %>"></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Sequence:<BR><INPUT TYPE=TEXT <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS=LABEL NAME=SEQUENCE SIZE=5 MAXLENGTH=10 VALUE="<%= SEQUENCE %>" OnChange="setDirty()" OnKeyPress="setDirty()"></TD>
<TD CLASS=LABEL>Retry Count:<BR><INPUT TYPE=TEXT <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS=LABEL NAME=RETRY_COUNT SIZE=10 MAXLENGTH=10 VALUE="<%= RETRY_COUNT %>" OnChange="setDirty()" OnKeyPress="setDirty()"></TD>
<TD CLASS=LABEL>Retry Wait Time:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=RETRY_WAIT_TIME SIZE=10 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> MAXLENGTH=10 VALUE="<%= RETRY_WAIT_TIME %>" OnChange="setDirty()" OnKeyPress="setDirty()"></TD>
<TD CLASS=LABEL COLSPAN=2>Transmission Type:<BR>
<SELECT NAME=TRANSMISSION_TYPE_ID CLASS=LABEL OnChange="checkFileName()" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>

<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = SQL & "SELECT * FROM TRANSMISSION_TYPE"
	Set RS = Conn.Execute(SQL)
	Do While Not RS.EOF
%>
<OPTION VALUE="<%= RS("TRANSMISSION_TYPE_ID") %>"><%= RS("NAME") %>
<%
RS.MoveNext
Loop
RS.Close
%>
</SELECT>
</TD>
</TR>
</TABLE>

<div style="visibility: hidden" id="XMLTemplate">
<table class="Label" ID="Table1" CELLSPACING="3" CELLPADDING="3">
<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">XML Template Filename (*Optional):<br></td>
</tr>
<tr>
<TD CLASS=LABEL COLSPAN=3>Filename (*Can be a run-time eval expression, like $IIF(..., ..., ...)):<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=XMLFILE SIZE=60 MAXLENGTH=2000 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> OnChange="setDirty()" OnKeyPress="setDirty()" VALUE="<%= ReplaceQuotesInText(XMLFILE) %>"></TD>
</tr>
<tr>
<TD CLASS=LABEL COLSPAN=3>XML Template Filename Description:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=XMLDESCRIPTION SIZE=60 MAXLENGTH=255 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> OnChange="setDirty()" OnKeyPress="setDirty()" VALUE="<%= ReplaceQuotesInText(XMLDESCRIPTION) %>"></TD>
</tr>
</table>
</div>

<div style="visibility: visible" id="RuleDiv">
<table class="Label" ID="Table2" CELLSPACING="3" CELLPADDING="3">
<tr>
<td CLASS="LABEL" VALIGN="BOTTOM">Output Filename Rule:</td>
</tr>
<tr>
<td>
<img NAME="BtnAttachRule" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
<img NAME="BtnDetachRule" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
</td>
<td width="250" nowrap>Rule Text:&nbsp;<span ID="RULE_TEXT" CLASS="LABEL" TITLE="<%=RSRULE_TEXT%>"><%=TruncateText(RSRULE_TEXT,30)%></span></td>
<td>Rule ID:&nbsp;<span ID="RULE_ID" CLASS="LABEL"><%=RULE_ID%></span></td>
</tr>
<tr>
<TD CLASS=LABEL COLSPAN=3>Filename Rule Description:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=DESCRIPTION SIZE=60 MAXLENGTH=255 <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> OnChange="setDirty()" OnKeyPress="setDirty()" VALUE="<%= ReplaceQuotesInText(DESCRIPTION) %>" ID="DESCRIPTION"></TD>
</tr>
</table>
</div>

<BR>
</FORM>
</BODY>
</HTML>
