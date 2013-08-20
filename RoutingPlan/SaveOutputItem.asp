<!--#include file="..\lib\common.inc"-->
<%

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
Function Swap(InData)
	If InData = "on" THen
		Swap = "Y"
	Else
		Swap = "N"
	End If
End Function
	
Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function

If Request.QueryString("STATUS") = "NEW" Then
SQL = ""
SQL = SQL & "INSERT INTO OUTPUT_ITEM ( OUTPUT_ITEM_ID, TRANSMISSION_SEQ_STEP_ID, "
SQL = SQL & "OUTPUTDEF_ID, SEQUENCE, RULE_ID ) VALUES ( "
SQL = SQL & NextPkey("OUTPUT_ITEM", "OUTPUT_ITEM_ID") & ", "
SQL = SQL & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", "
SQL = SQL & Request.Form("OUTPUTDEF_ID") & ", "
SQL = SQL & Request.Form("SEQUENCE") & ", "
if IsNull(Request.Form("ENABLERULE_ID")) or Request.Form("ENABLERULE_ID") = "" then
	SQL = SQL & "NULL )"
else
	SQL = SQL & Request.Form("ENABLERULE_ID") & ") "
End if
set rs = conn.Execute(SQL)
End If 

If Request.QueryString("STATUS") = "UPDATE" Then
If Request.Form("OUTPUT_ITEM_ID") <> "" Then
SQL = ""
SQL = SQL & "UPDATE OUTPUT_ITEM SET "
SQL = SQL & "OUTPUTDEF_ID=" & Request.Form("OUTPUTDEF_ID") & ", "
SQL = SQL & "SEQUENCE=" & Request.Form("SEQUENCE") & ", "
if IsNull(Request.Form("ENABLERULE_ID")) or Request.Form("ENABLERULE_ID") = "" then
	SQL = SQL & "RULE_ID=NULL"
else
	SQL = SQL & "RULE_ID=" & Request.Form("ENABLERULE_ID")
End if
SQL = SQL & " WHERE OUTPUT_ITEM_ID=" & Request.Form("OUTPUT_ITEM_ID")
set rs5 = conn.Execute(SQL)
End If
End If

%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	parent.frames("TOP").document.all.StatusSpan.InnerHtml = "Output Item Saved"
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>