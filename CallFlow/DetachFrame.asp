<!--#include file="..\lib\common.inc"-->
<%
If Request.QueryString("FRAMEID") <> "" AND Request.QueryString("CFID") <> "" Then

	SQL = ""
	SQL = SQL & "DELETE FROM FRAME_ORDER WHERE "
	SQL = SQL & "CALLFLOW_ID=" & Request.QueryString("CFID") & " AND "
	SQL = SQL & "FRAME_ID=" & Request.QueryString("FRAMEID")
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	Set RS = Conn.Execute(SQL)
	
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	top.frames.location.href = "CallFlow-f.asp?CFID=<%= Request.QueryString("CFID") %>"
End Sub

-->
</SCRIPT>
</HEAD>
<% End If %>
<BODY>



</BODY>
</HTML>
