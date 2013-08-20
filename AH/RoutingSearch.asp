<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Search</TITLE>
<STYLE>
BODY 
{ 
	background:#d6cfbd;
	Font-Family:Verdana;
	Font-Size:10
}
</STYLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload

End Sub

Sub ExeSearch()
	document.all.FrmSearch.submit()
End Sub

Sub ClearSearch
	document.all.LOB_CD.Value = "" 
	document.all.STATE.Value = "" 
	document.all.DESTINATION_TYPE.Value = "" 
	document.all.INPUT_SYSTEM_NAME.Value = "" 
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
<FORM NAME=FrmSearch ACTION="RoutingSearchResults.asp" METHOD=POST TARGET=WORKAREA>
<TABLE>
<TR>
<TD CLASS=LABEL VALIGN=BOTTOM>Destination Type:<BR><INPUT CLASS=LABEL TYPE=TEXT STYLE="TEXT-TRANSFORM:UPPERCASE" NAME=DESTINATION_TYPE></TD>
<TD CLASS=LABEL VALIGN=BOTTOM>Input System Name:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=INPUT_SYSTEM_NAME STYLE="TEXT-TRANSFORM:UPPERCASE"></TD>
<TD CLASS=LABEL>State:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=STATE STYLE="TEXT-TRANSFORM:UPPERCASE" MAXLENGTH=2 SIZE=2></TD>
<TD CLASS=LABEL>LOB:<BR>
<SELECT NAME=LOB_CD CLASS=LABEL STYLE="WIDTH:75">
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT LOB_CD FROM LOB WHERE LOB_CD IS NOT NULL"
	Set RS = Conn.Execute(SQLST)
Do While Not RS.EOF
%>
<OPTION VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_CD") %>
<%
RS.MoveNext
Loop
RS.CLose
%>
</SELECT></TD>
</TR>
<TR>
<TD CLASS=LABEL></TD>
<TD CLASS=LABEL></TD>
<TD CLASS=LABEL></TD>
<TD CLASS=LABEL></TD>
</TR>
</TABLE>
</BODY>
</HTML>

