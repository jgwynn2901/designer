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
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FORM NAME="FrmSave" ACTION="Alt_Name_Save.asp" TARGET="HIDDENPAGE" METHOD="POST">
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
<TD COLSPAN=2 CLASS=LABEL>Alternate Name:<BR><INPUT CLASS=LABEL NAME=ALT_NAME TYPE=TEXT SIZE=80 MAXLENGTH=80 ></TD>
</TR>
</TABLE>

</FORM>
</BODY>
</HTML>
