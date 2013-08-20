<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<STYLE>
BODY { 
		background:#d6cfbd;
		Font-Family:Verdana;
		Font-Size:10 
		}
</STYLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload

End Sub

Sub ClearSearch
	document.all.NAME.value = ""
	document.all.LOB_CD.value = ""
	document.all.DESCRIPTION.value = ""
End Sub

Sub ExeSearch()
	If document.all.NAME.value = "" AND document.all.LOB_CD.value = "" AND document.all.DESCRIPTION.value = "" Then
		MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	Else
		FrmSearch.submit
	End If
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
<FORM NAME=FrmSearch ACTION="CallFlowSearchResults.asp" METHOD=POST TARGET=WORKAREA>
<TABLE>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT NAME=NAME CLASS=LABEL STYLE="TEXT-TRANSFORM:UPPERCASE"></TD>
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
<TD CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT NAME=DESCRIPTION CLASS=LABEL STYLE="TEXT-TRANSFORM:UPPERCASE"></TD>
<TD VALIGN=BOTTOM></TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
