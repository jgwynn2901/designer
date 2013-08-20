<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<STYLE>
BODY { 
		background:#d6cfbd;
		Font-Family:Verdana;
		Font-Size:10 
		}
</STYLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub ClearSearch()
	document.all.POLICY_NUMBER.value = ""
	document.all.POLICY_DESC.value = ""
	document.all.LOB_CD.value = ""
End Sub

Sub ExeSearch()
If document.all.POLICY_NUMBER.value = "" AND document.all.POLICY_DESC.value = "" AND document.all.EXPIRATION_DATE.value = "" AND document.all.EFFECTIVE_DATE.value = "" AND document.all.LOB_CD.value = "" Then
	MsgBox "Please enter search criteria!", 0 , "FNSNetDesigner"
Else
	FrmSearch.submit
End If
End Sub

Sub window_onload
	'document.all.POLICY_NUMBER.focus()
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0>
<FORM Name="FrmSearch" TARGET="WORKAREA" METHOD=POST ACTION="PolicySearchResults.asp">
<TABLE>
<TR>
<TD CLASS=LABEL>Policy number:<BR><INPUT TYPE=TEXT NAME=POLICY_NUMBER CLASS=LABEL STYLE="TEXT-TRANSFORM:UPPERCASE"></TD>
<TD CLASS=LABEL>Policy desc:<BR><INPUT TYPE=TEXT NAME=POLICY_DESC CLASS=LABEL STYLE="TEXT-TRANSFORM:UPPERCASE"></TD>
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
<TD CLASS=LABEL></TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
