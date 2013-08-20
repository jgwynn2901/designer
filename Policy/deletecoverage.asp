<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<%
If  Request.QueryString("COVID") <> "" Then
SQLDEL = ""
SQLDEL = SQLDEL & "DELETE FROM COVERAGE WHERE COVERAGE_ID=" & Request.QueryString("COVID")
	Set RSDelete = Conn.Execute(SQLDEL)
	strError = CheckADOErrors(Conn,"Coverage " & ACTION)
End If
%>