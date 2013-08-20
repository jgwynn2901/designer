<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Search</TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

<!--#include file="..\lib\Help.asp"-->

Sub window_onload
	document.all.SearchType(0).checked = True
End Sub

Sub BtnClear_OnClick
	document.all.LOB_CD.Value = ""
	document.all.NAME.Value = ""  
	document.all.DESCRIPTION.Value = "" 
	document.all.CALLFLOW_ID.Value = ""
	document.all.FRAME_NAME.Value = ""
End Sub

Sub BtnSearch_OnCLick
'If document.all.CALLFLOW_ID.Value = "" AND document.all.NAME.Value = "" AND document.all.DESCRIPTION.Value = "" AND document.all.LOB_CD.Value = "" And document.all.FRAME_NAME.Value = "" Then
'	MsgBox "Enter criteria and choose Search", 0, "FNSDesigner"
'Else
	document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
	document.all.FrmSearch.Submit()
'End If
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub
-->
</SCRIPT>
</HEAD>
<BODY  rightmargin=0 leftmargin=0 bottommargin=0 topmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<FORM NAME=FrmSearch ACTION="CallFlowResults.asp" METHOD=POST TARGET=WORKAREA>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Call Flow Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="ahsid" VALUE="<%= Request.QueryString("ahsid") %>">
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label"  cellspacing=0 cellpadding=0>
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" Title="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>
<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 WIDTH="100%">
<TR WIDTH=1><TD>
<TABLE CELLPADDING=0 CELLSPACING=0>
<TR>
<TD CLASS=LABEL COLSPAN=2 VALIGN=BOTTOM>Name:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=NAME SIZE=38></TD>
</TD>
<TD CLASS=LABEL COLSPAN=2>LOB:<BR>
<SELECT NAME=LOB_CD CLASS=LABEL>
<OPTION VALUE="">
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
	Set RS = Conn.Execute(SQLST)
Do While Not RS.EOF
%>
<OPTION VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_NAME") %>
<%
RS.MoveNext
Loop
RS.CLose
%>
</SELECT></TD>
</TR>
</TABLE>
<TABLE CELLPADDING=0 CELLSPACING=0>
<TR>
<TD CLASS=LABEL>CFID:<BR><INPUT TYPE=TEXT SIZE=5 CLASS=LABEL NAME=CALLFLOW_ID></TD>
<TD CLASS=LABEL COLSPAN=3>Description:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=DESCRIPTION SIZE=55></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=3>Frame Name:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=FRAME_NAME SIZE=68></TD>
</TR>
</TABLE>
<TABLE>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</TABLE>
</TD><TD VALIGN=TOP ALIGN=RIGHt>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON ACCESSKEY="C" NAME=BtnSearch>Sear<U>c</U>h</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON ACCESSKEY="L" NAME=BtnClear>C<U>l</U>ear</BUTTON></TD>
</TR>
</TABLE>
</TD></TR></TABLE>
</BODY>
</HTML>

