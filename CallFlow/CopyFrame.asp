<!--#include file="..\lib\common.inc"-->
<%
Response.Expires=0

If Request.QueryString("FRAMEID") <> "" AND Request.QueryString("CFID") <> "" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = "{call Designer.CopyFrame(" &  Request.QueryString("FRAMEID") & " ,{resultset 1, outFrameId, StatusMsg, StatusNum})}"
	Set RSCopy = Conn.Execute(SQL)
	SQL = ""
	SQL = SQL & "{call Designer.CopyFrameOrder(" &  Request.QueryString("FRAMEID") & ", "
	SQL = SQL & Request.QueryString("CFID") & "," 
	SQL = SQL & RSCopy("outFrameId") & ","
	SQL = SQL & Request.QueryString("CFID") & ",{resultset 1, StatusMsg, StatusNum})}"
	Set RS=Conn.Execute(SQL)
	SQLDelete = ""
	SQLDelete = SQLDelete & "DELETE FROM FRAME_ORDER WHERE CALLFLOW_ID=" & Request.QueryString("CFID") & " AND "
	SQLDelete = SQLDelete & "FRAME_ID=" & Request.QueryString("FRAMEID")
	Set RSDel = Conn.Execute(SQLDelete)

%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	top.frames.location.href = "CallFlow-f.asp?CFID=<%= Request.QueryString("CFID") %>&FRAMEID=<%= RSCopy("outFrameId") %>"
End Sub
-->
</SCRIPT>
</HEAD>
<% End If %>
<BODY>



</BODY>
</HTML>
