<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\validvalues.inc"-->
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
	document.all.STATE.Value = "" 
	document.all.DESCRIPTION.Value = "" 
	document.all.DESTINATION_TYPE.Value = "" 
	document.all.INPUT_SYSTEM_NAME.Value = "" 
	document.all.ROUTING_PLAN_ID.Value = ""
	document.all.s_EnabledFlag.Value = ""
End Sub

Sub BtnSearch_OnCLick
'If document.all.LOB_CD.Value = "" AND document.all.ROUTING_PLAN_ID.Value = "" AND document.all.STATE.Value = "" AND document.all.DESCRIPTION.Value = "" AND document.all.DESTINATION_TYPE.Value = "" AND document.all.INPUT_SYSTEM_NAME.Value = "" Then
'	MsgBox "Enter criteria and choose Search", 0, "FNSDesigner"
'Else
	document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
	document.all.FrmSearch.Submit()
'End IF
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub
-->
</SCRIPT>
</HEAD>
<BODY  rightmargin=0 leftmargin=0 bottommargin=0 topmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<FORM NAME=FrmSearch ACTION="RoutingPlanResults.asp" METHOD=POST TARGET=WORKAREA>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Routing Plan Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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
<TABLE BORDER=0 WIDTH=100%><TR><TD>
<TABLE>
<TR>
<TD CLASS=LABEL COLSPAN=2 VALIGN=BOTTOM>Destination Type:<BR>
<SELECT NAME=DESTINATION_TYPE CLASS=LABEL>
<%= GetValidValuesHTML("DESTINATION_TYPE", "", true) %>
</SELECT></TD>
<TD CLASS=LABEL VALIGN=BOTTOM><NOBR>Input System Name:<BR>
<SELECT NAME="INPUT_SYSTEM_NAME" CLASS=LABEL STYLE="WIDTH:100%">
<OPTION VALUE="">
<OPTION VALUE="FNS NET">FNS NET
<OPTION VALUE="OPEN BASIC">OPEN BASIC
<OPTION VALUE="FNSINETP1">FNSINETP1
</SELECT>

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
<TD CLASS=LABEL VALIGN=BOTTOM><NOBR>Enabled:<BR>
	<SELECT NAME="s_EnabledFlag" CLASS=LABEL STYLE="WIDTH:100%">
		<OPTION VALUE="">
		<OPTION VALUE="Y">Enabled
		<OPTION VALUE="N">Disabled
	</SELECT>
</TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>RPID:<BR><INPUT TYPE=TEXT SIZE=5 CLASS=LABEL NAME=ROUTING_PLAN_ID></TD>
<TD CLASS=LABEL COLSPAN=3>Description:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=DESCRIPTION SIZE=60></TD>
<TD CLASS=LABEL>State:<BR>
<SELECT NAME=STATE CLASS=LABEL>
<OPTION VALUE="">
<!--#include file="..\lib\states.asp"-->
</SELECT>
</TD>
</TR>
</TABLE>
<TABLE BORDER=0>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
<td width = 60></td>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON ACCESSKEY="C" NAME=BtnSearch ID="Button1">Sear<U>c</U>h</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON ACCESSKEY="L" NAME=BtnClear ID="Button2">C<U>l</U>ear</BUTTON></TD>
</tr>
</TABLE>
</TD><TD VALIGN=TOP ALIGN=RIGHT>

</TD></TR></TABLE>
</BODY>
</HTML>

