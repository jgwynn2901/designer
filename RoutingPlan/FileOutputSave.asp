<!--#include file="..\lib\common.inc"-->
<%

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function

If Request.QueryString("STATUS") = "NEW" Then

SQL = "INSERT INTO OUTPUT_FILE (OUTPUT_FILE_ID, OUTPUTDEF_ID, OUTPUT_FILE_NAME, OUTPUT_FILE_FORMAT, REPORT_FILE_NAME, ENABLE_RULE_ID) " 
SQL = SQL & "VALUES ( "
SQL = SQL & NextPkey("OUTPUT_FILE", "OUTPUT_FILE_ID") & ", "
SQL = SQL & Request.Form("OUTPUTDEF_ID") & ", '"
SQL = SQL & Request.Form("OUTPUT_FILE") & "', '"
SQL = SQL & Request.Form("OUTPUT_FORMAT") & "', '"
SQL = SQL & Request.Form("CR_FILE") & "', "
SQL = SQL & Request.Form("ENABLING_RULE_ID") & ")"
set rs = conn.Execute(SQL)
Conn.Close 
set RS=nothing
set Conn=nothing
End If 

If Request.QueryString("STATUS") = "UPDATE" Then
If Request.Form("OUTPUT_FILE_ID") <> "" Then
SQL = "UPDATE OUTPUT_FILE SET "
SQL = SQL & "OUTPUTDEF_ID=" & Request.Form("OUTPUTDEF_ID") & ", "
SQL = SQL & "OUTPUT_FILE_FORMAT='" & Request.Form("OUTPUT_FORMAT") & "', "
SQL = SQL & "REPORT_FILE_NAME='" & Request.Form("CR_FILE") & "', "
SQL = SQL & "ENABLE_RULE_ID=" & Request.Form("ENABLING_RULE_ID") & ", "
SQL = SQL & "OUTPUT_FILE_NAME='" & Request.Form("OUTPUT_FILE")
SQL = SQL & "' WHERE OUTPUT_FILE_ID=" & Request.Form("OUTPUT_FILE_ID")
set rs = conn.Execute(SQL)
Conn.Close 
set RS=nothing
set Conn=nothing
End If
End If

If Request.QueryString("STATUS") = "DELETE" Then
If Request.QueryString("OFID") <> "" Then
SQL = "DELETE FROM OUTPUT_FILE WHERE OUTPUT_FILE_ID=" & Request.QueryString("OFID")
set rs = conn.Execute(SQL)
Conn.Close 
set RS=nothing
set Conn=nothing
End If
End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	parent.frames("TOP").document.all.StatusSpan.InnerHtml = "File Output Saved"
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>