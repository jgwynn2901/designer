<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
%>

<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub ExeSave()
	FrmSave.target = "HIDDENPAGE"
	FrmSave.submit()
	SpanStatus.innerHTML = "Update Successful!"
End Sub
-->
</SCRIPT>
</HEAD>
<%
dim oRS, cSQL
dim cName, cType, cTitle, cPhone, cFax, cEMail, cDesc

cName = ""
cType = ""
cTitle = ""
cPhone = ""
cFax = ""
cEMail = ""
cDesc = ""
if Request.QueryString("EDIT") <> "" then
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.MaxRecords = MAXRECORDCOUNT
	cSQL = "SELECT * FROM CONTACT WHERE CONTACT_ID=" & Request.QueryString("EDIT")
	oRS.Open cSQL, CONNECT_STRING, adOpenStatic, adLockReadOnly, adCmdText
	cName = trim(oRS("NAME"))
	cType = trim(oRS("TYPE"))
	cTitle = trim(oRS("TITLE"))
	cPhone = trim(oRS("PHONE"))
	cFax = trim(oRS("FAX"))
	cEMail = trim(oRS("EMAIL"))
	cDesc = trim(oRS("DESCRIPTION"))
	oRS.Close
	set oRS = nothing
end if
%>

<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FORM NAME="FrmSave" ACTION="Contact_Save.asp" TARGET="HIDDENPAGE" METHOD="POST">
<input Type="Hidden" Name="ContactID" Value="<%=Request.QueryString("EDIT")%>">
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>
<INPUT TYPE=HIDDEN NAME=AHSID VALUE="<%= Request.QueryString("AHSID") %>">
<TABLE>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT CLASS=LABEL NAME=CNT_NAME VALUE="<%=cName%>" TYPE=TEXT SIZE=40 MAXLENGTH=60 ></TD>
</TR>
<TR>
<TD CLASS=LABEL>Type:<BR><INPUT CLASS=LABEL NAME=CNT_TYPE VALUE="<%=cType%>" TYPE=TEXT SIZE=40 MAXLENGTH=40 ></TD>
<TD CLASS=LABEL>Title:<BR><INPUT CLASS=LABEL NAME=CNT_TITLE VALUE="<%=cTitle%>" TYPE=TEXT SIZE=40 MAXLENGTH=80 ></TD>
</TR>
<TR>
<TD CLASS=LABEL>Phone:<BR><INPUT CLASS=LABEL NAME=CNT_PHONE VALUE="<%=cPhone%>" TYPE=TEXT SIZE=14 MAXLENGTH=14 ></TD>
<TD CLASS=LABEL>Fax:<BR><INPUT CLASS=LABEL NAME=CNT_FAX VALUE="<%=cFax%>" TYPE=TEXT SIZE=10 MAXLENGTH=10 ></TD>
</TR>
<TR>
<TD CLASS=LABEL>E-Mail:<BR><INPUT CLASS=LABEL NAME=CNT_EMAIL VALUE="<%=cEMail%>" TYPE=TEXT SIZE=40 MAXLENGTH=255 ></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Description:<BR><INPUT CLASS=LABEL NAME=CNT_DESC VALUE="<%=cDesc%>" TYPE=TEXT SIZE=85 MAXLENGTH=2000 ></TD>
</TR>
</TABLE>

</FORM>
</BODY>
</HTML>
