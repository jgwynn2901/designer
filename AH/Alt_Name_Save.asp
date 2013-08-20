<!--#include file="..\lib\common.inc"-->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
ConnectionString = CONNECT_STRING
Conn.Open ConnectionString
	
QSQL = ""
QSQL = QSQL & "{call Designer.GetValidSeq('ALTERNATE_NAME', 'ALTERNATE_NAME_ID', {resultset 1, outResult})}"
Set RSNextID = Conn.Execute(QSQL)

SQL = ""
SQL = SQL & "INSERT INTO ALTERNATE_NAME (ALTERNATE_NAME_ID, ACCNT_HRCY_STEP_ID, NAME) VALUES ("
SQL = SQL & RSNextID("outResult") & ", "
SQL = SQL & Request.Form("AHSID") & ", "
SQL = SQL & "'" & Replace(Request.Form("ALT_NAME"), "'", "''") & "') "
Set RS = conn.Execute(SQL)

RSNextID.Close
Conn.Close
%>
<HTML>
<HEAD>

</HEAD>
<BODY>



</BODY>
</HTML>
