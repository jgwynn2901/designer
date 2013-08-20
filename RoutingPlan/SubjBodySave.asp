<!--#include file="..\lib\common.inc"-->
<%

Set Conn = Server.CreateObject("ADODB.Connection")
ConnectionString = CONNECT_STRING
Conn.Open ConnectionString

If Request.QueryString("STATUS") = "NEW" Then

SQL = "INSERT INTO OUTPUT_SUBJECT_BODY (TRANSMISSION_SEQ_STEP_ID, SUBJECT_FILE_NAME, BODY_FILE_NAME) " 
SQL = SQL & "VALUES ( "
SQL = SQL & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", '"
SQL = SQL & Request.Form("SUBJECT_FILE") & "', '"
SQL = SQL & Request.Form("BODY_FILE") & "')"
set rs = conn.Execute(SQL)
Conn.Close 
set RS=nothing
set Conn=nothing
End If 

If Request.QueryString("STATUS") = "UPDATE" Then
If Request.Form("OUTPUT_SUBJECT_BODY_ID") <> "" Then
SQL = "UPDATE OUTPUT_SUBJECT_BODY SET "
SQL = SQL & "TRANSMISSION_SEQ_STEP_ID=" & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", "
SQL = SQL & "SUBJECT_FILE_NAME='" & Request.Form("SUBJECT_FILE") & "', "
SQL = SQL & "BODY_FILE_NAME='" & Request.Form("BODY_FILE")
SQL = SQL & "' WHERE OUTPUT_SUBJECT_BODY_ID=" & Request.Form("OUTPUT_SUBJECT_BODY_ID")
set rs = conn.Execute(SQL)
Conn.Close 
set RS=nothing
set Conn=nothing
End If
End If

If Request.QueryString("STATUS") = "DELETE" Then
If Request.QueryString("OSBID") <> "" Then
SQL = "DELETE FROM OUTPUT_SUBJECT_BODY WHERE OUTPUT_SUBJECT_BODY_ID=" & Request.QueryString("OSBID")
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
	parent.frames("TOP").document.all.StatusSpan.InnerHtml = "Email Subject/Body template file names are saved"
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>