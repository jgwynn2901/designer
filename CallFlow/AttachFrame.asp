<!--#include file="..\lib\common.inc"-->
<%
If Request.QueryString("FRAMEID") <> "" AND Request.QueryString("CFID") <> "" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
	Dim I, NextSequence
	
	If Request.QueryString("SEQUENCE") <> "" Then
		I = 1
		NextSequence = Clng(Request.QueryString("SEQUENCE")) + 1
		SQLCheck = ""
		SQLCheck = SQLCheck & "SELECT SEQUENCE FROM FRAME_ORDER WHERE CALLFLOW_ID=" & Request.QueryString("CFID")
		Set RSCheck = Conn.Execute(SQLCheck)
	End If

	SQL2 = ""
	SQL2 = SQL2 & "INSERT INTO FRAME_ORDER (FRAME_ID, CALLFLOW_ID" 
	SQL2 = SQL2 & ", SEQUENCE "
	SQL2 = SQL2 & ",TITLE, ATTRIBUTE_PREFIX, ENABLEDRULE_ID, VALIDRULE_ID, "
	SQL2 = SQL2 & "MODAL_FLG, ENTRY_ACTION_ID, ACTION_ID, HELPSTRING, "
	SQL2 = SQL2 & "DESCRIPTION, TYPE, SQLSELECT, SQLFROM, SQLWHERE, SQLORDERBY, "
	SQL2 = SQL2 & "MAXPAGERESULTROWS, ONEROWAUTOSELECT_FLG "
	SQL2 = SQL2 & ") VALUES ("
	SQL2 = SQL2 & Request.QueryString("FRAMEID") & ", " 
	SQL2 = SQL2 & Request.QueryString("CFID") 
	SQL2 = SQL2 & ",0"
	SQL2 = SQL2 & ",'-999999999', '-999999999', -999999999, -999999999, "
	SQL2 = SQL2 & "'U', -999999999, -999999999, '-999999999', "
	SQL2 = SQL2 & "'-999999999', '-999999999', '-999999999', '-999999999', '-999999999', '-999999999', "
	SQL2 = SQL2 & "-999999999, 'U' "
	SQL2 = SQL2 & ")" 
	Set RS = Conn.Execute(SQL2)
	
%>
<HTML>
<HEAD>
<SCRIPT LANGUAGE="VBScript">
<!--
	top.frames.location.href = "CallFlow-f.asp?CFID=<%= Request.QueryString("CFID") %>"
-->
</SCRIPT>
</HEAD>
<% End IF %>
<BODY>



</BODY>
</HTML>
