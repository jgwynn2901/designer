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
	
Function SwapNull(InData)
	If InData = "" OR isNull(InData) Then
		SwapNull="null"
	Else
		SwapNull = InData
	End If
End function
	
Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function


If Request.QueryString("STATUS") = "NEW" Then
nTRANSMISSION_SEQ_STEP_ID = NextPKey("TRANSMISSION_SEQ_STEP", "TRANSMISSION_SEQ_STEP_ID")
SQL = ""
SQL = SQL & "INSERT INTO TRANSMISSION_SEQ_STEP ( "
SQL = SQL & "TRANSMISSION_SEQ_STEP_ID, ROUTING_PLAN_ID, "
SQL = SQL & "SEQUENCE, RETRY_COUNT, RETRY_WAIT_TIME, "
SQL = SQL & "DESTINATION_STRING, TRANSMISSION_TYPE_ID, BATCH_HOLD, ALT_DESTINATION_STRING ) VALUES ("
SQL = SQL & nTRANSMISSION_SEQ_STEP_ID & ", "
SQL = SQL & Request.Form("ROUTING_PLAN_ID") & ", "
SQL = SQL & Request.Form("SEQUENCE") & ", "
SQL = SQL &  "'" & Request.Form("RETRY_COUNT") & "', "
SQL = SQL &  "'" & Request.Form("RETRY_WAIT_TIME") & "', "
SQL = SQL & "'" & Replace(Replace(Request.Form("DESTINATION_STRING"),"""", ""),"'","''") & "', "
SQL = SQL & Request.Form("TRANSMISSION_TYPE_ID") & ", "
SQL = SQL & "'" & Swap(Request.Form("BATCH_HOLD")) & "', "
SQL = SQL & "'" & Replace(Request.Form("ALT_DESTINATION_STRING"),"""", "") & "') "

set rs = conn.Execute(SQL)
if Request.Form("FILENAME_FLG") = "Y"  and Request.Form("FILENAME_RULE_ID") <> "" then
	SQL = ""
	SQL = SQL & "INSERT INTO OUTPUT_FILENAME (TRANSMISSION_SEQ_STEP_ID, RULE_ID, DESCRIPTION) VALUES ("
	SQL = SQL & nTRANSMISSION_SEQ_STEP_ID & ", "
	SQL = SQL & Request.Form("FILENAME_RULE_ID") & ", "
	SQL = SQL & "'" & Request.Form("DESCRIPTION") & "') "
	set rs = conn.Execute(SQL)
end if

if Request.Form("XML_FILENAME_FLG") = "Y"  and Request.Form("XMLFILE") <> "" then
	SQL = ""
	SQL = SQL & "INSERT INTO OUTPUT_XMLTEMPLATE (TRANSMISSION_SEQ_STEP_ID, FILE_NAME, DESCRIPTION) VALUES ("
	SQL = SQL & nTRANSMISSION_SEQ_STEP_ID & ", "
	SQL = SQL & "'" & Request.Form("XMLFILE") & "', "
	SQL = SQL & "'" & Request.Form("XMLDESCRIPTION") & "') "
	set rs = conn.Execute(SQL)
end if



'------------------------JBOR-0055-------------------------------'
  
 nRPID = Request.Form("ROUTING_PLAN_ID")
 nEDIOutboundItemID = NextPKey("EDI_OUTBOUND_ITEM", "EDI_OUTBOUND_ITEM_ID")
 nChar = chr(13) & chr(10)
  
 if  TRIM(Request.Form("TRANSMISSION_TYPE_ID")) = "5" then
 
		
     cSQL = "select DESCRIPTION from  ROUTING_PLAN where ROUTING_PLAN_ID = " & nRPID 
	 set rs = conn.Execute(cSQL)
	 with rs
	   cDesc = .Fields(0)
	 end with
		 
	 cSQLTransmission = "select TRANSMISSION_SEQ_STEP_ID from  TRANSMISSION_SEQ_STEP where ROUTING_PLAN_ID = " & nRPID
     set rs = conn.Execute(cSQLTransmission)
     with rs
	   nTransmissionSeqStepId = .Fields(0)
	 end with
	 
	 Conn.BeginTrans
     
		cSP1 = "INSERT INTO EDI_OUTBOUND_ITEM (EDI_OUTBOUND_ITEM_ID,TRANSMISSION_SEQ_STEP_ID,NAME,DESCRIPTION,DELIMITER,RECORD_DELIMITER) VALUES ('"& nEDIOutboundItemID &"','"& nTRANSMISSION_SEQ_STEP_ID &"','"& cDesc &"','',',','"& nChar &"')"
		Conn.Execute (cSP1)

		cSP2 = "INSERT INTO EDI_OUTBOUND_TOP_SEGMENTS( EDI_OUTBOUND_ITEM_ID,EDI_OUTBOUND_SEGMENT_ID) VALUES ('"& nEDIOutboundItemID &"','1')"
		Conn.Execute (cSP2)

     Conn.CommitTrans  
     
end if
'------------------------JBOR-0055-------------------------------'	

		
End If


If Request.QueryString("STATUS") = "UPDATE" Then
	SQl2 = ""
	SQL2 = SQL2 & "UPDATE TRANSMISSION_SEQ_STEP SET "
	SQL2 = SQL2 & "SEQUENCE=" & Request.Form("SEQUENCE") & ", "
	SQL2 = SQL2 & "RETRY_COUNT=" & Request.Form("RETRY_COUNT") & ", "
	SQL2 = SQL2 & "RETRY_WAIT_TIME=" & Request.Form("RETRY_WAIT_TIME") & ", "
	SQL2 = SQL2 & "DESTINATION_STRING='" & Replace(Request.Form("DESTINATION_STRING"),"'","''") & "', "
	SQL2 = SQL2 & "TRANSMISSION_TYPE_ID=" & Request.Form("TRANSMISSION_TYPE_ID") & ", "
	SQL2 = SQL2 & "BATCH_HOLD='" & Swap(Request.Form("BATCH_HOLD")) & "', "
	SQL2 = SQL2 & "ALT_DESTINATION_STRING='" & Request.Form("ALT_DESTINATION_STRING") & "' "
	SQL2 = SQL2 & "WHERE TRANSMISSION_SEQ_STEP_ID ="& Request.Form("TRANSMISSION_SEQ_STEP_ID") 
	set rs = conn.Execute(SQL2)

	if Request.Form("FILENAME_FLG") = "Y" and Request.Form("FILENAME_RULE_ID") <> ""  then
		SQL2 = ""
		SQl2 = SQl2 & "SELECT OUTPUT_FILENAME_ID FROM OUTPUT_FILENAME WHERE "
		SQl2 = SQl2 & "TRANSMISSION_SEQ_STEP_ID = " & Request.Form("TRANSMISSION_SEQ_STEP_ID")
		set rs = conn.Execute(SQL2)
		if rs.eof then
			set rs = nothing
			SQl2 = ""
			SQl2 = SQl2 & "INSERT INTO OUTPUT_FILENAME (TRANSMISSION_SEQ_STEP_ID, RULE_ID, DESCRIPTION) VALUES ("
			SQl2 = SQl2 & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", "
			SQl2 = SQl2 & Request.Form("FILENAME_RULE_ID") & ", "
			SQl2 = SQl2 & "'" & Request.Form("DESCRIPTION") & "') "
			set rs = conn.Execute(SQL2)
		else
			SQl2 = ""
			SQL2 = SQL2 & "UPDATE OUTPUT_FILENAME SET "
			SQL2 = SQL2 & "RULE_ID=" & Request.Form("FILENAME_RULE_ID") & ", "
			SQL2 = SQL2 & "TRANSMISSION_SEQ_STEP_ID=" & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", "
			SQL2 = SQL2 & "DESCRIPTION=" & " '" & Request.Form("DESCRIPTION") & "' "
			SQL2 = SQL2 & "WHERE OUTPUT_FILENAME_ID ="& rs("OUTPUT_FILENAME_ID") 
			set rs = nothing
			set rs = conn.Execute(SQL2)
		end if
		set rs = nothing
	else
		if Request.Form("TRANSMISSION_SEQ_STEP_ID") <> "" then
			SQLD = ""
			SQLD = SQLD & "DELETE FROM OUTPUT_FILENAME WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.Form("TRANSMISSION_SEQ_STEP_ID")
			set rs = conn.Execute(SQLD)
		end if
	end if
	
	if Request.Form("XML_FILENAME_FLG") = "Y" and Request.Form("XMLFILE") <> ""  then
		SQL2 = ""
		SQl2 = SQl2 & "SELECT OUTPUT_XMLTEMPLATE_ID FROM OUTPUT_XMLTEMPLATE WHERE "
		SQl2 = SQl2 & "TRANSMISSION_SEQ_STEP_ID = " & Request.Form("TRANSMISSION_SEQ_STEP_ID")
		set rs = conn.Execute(SQL2)
		if rs.eof then
			set rs = nothing
			SQL = ""
			SQL = SQL & "INSERT INTO OUTPUT_XMLTEMPLATE (TRANSMISSION_SEQ_STEP_ID, FILE_NAME, DESCRIPTION) VALUES ("
			SQL = SQL & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", "
			SQL = SQL & "'" & Request.Form("XMLFILE") & "', "
			SQL = SQL & "'" & Request.Form("XMLDESCRIPTION") & "') "
			set rs = conn.Execute(SQL)
		else
			SQl2 = ""
			SQL2 = SQL2 & "UPDATE OUTPUT_XMLTEMPLATE SET "
			SQL2 = SQL2 & "FILE_NAME='" & Request.Form("XMLFILE") & "', "
			SQL2 = SQL2 & "TRANSMISSION_SEQ_STEP_ID=" & Request.Form("TRANSMISSION_SEQ_STEP_ID") & ", "
			SQL2 = SQL2 & "DESCRIPTION=" & " '" & Request.Form("XMLDESCRIPTION") & "' "
			SQL2 = SQL2 & "WHERE OUTPUT_XMLTEMPLATE_ID ="& rs("OUTPUT_XMLTEMPLATE_ID") 
			set rs = nothing
			set rs = conn.Execute(SQL2)
		end if
		set rs = nothing
	else
		if Request.Form("TRANSMISSION_SEQ_STEP_ID") <> "" then
			SQLD = ""
			SQLD = SQLD & "DELETE FROM OUTPUT_XMLTEMPLATE WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.Form("TRANSMISSION_SEQ_STEP_ID")
			set rs = conn.Execute(SQLD)
		end if
	end if

End If

If Request.QueryString("STATUS") = "DELETE" Then
SQLD = ""
SQLD = SQLD & "DELETE FROM OUTPUT_FILENAME WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
set rs = conn.Execute(SQLD)

SQLD = ""
SQLD = SQLD & "DELETE FROM OUTPUT_XMLTEMPLATE WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
set rs = conn.Execute(SQLD)

'------------------------JBOR-0065-------------------------------'
			
SQL = ""
SQl = SQL & "SELECT * FROM EDI_OUTBOUND_ITEM WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
set rsEDI = conn.Execute(SQL)

Conn.BeginTrans
 
 	cSP2 = "DELETE FROM EDI_OUTBOUND_TOP_SEGMENTS WHERE EDI_OUTBOUND_ITEM_ID=" & rsEDI("EDI_OUTBOUND_ITEM_ID")
	Conn.Execute (cSP2)
	cSP1 = "DELETE FROM EDI_OUTBOUND_ITEM WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	Conn.Execute (cSP1)


 Conn.CommitTrans  
'------------------------JBOR-0065-------------------------------'	

SQLD = ""
SQLD = SQLD & "DELETE FROM TRANSMISSION_SEQ_STEP WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
set rs = conn.Execute(SQLD)
End If

If Request.QueryString("STATUS") = "DELETE_OUTPUT_ITEM" Then
SQL = ""
SQL = SQL & "DELETE FROM OUTPUT_ITEM WHERE OUTPUT_ITEM_ID=" & Request.QueryString("OUTPUT_ITEM_ID")
set rs = conn.Execute(SQL)
End If

If Request.QueryString("STATUS") = "COPY" Then

If Request.QueryString("TRANSMISSION_SEQ_STEP_ID") <> "" Then
	NextTrans = NextPKey("TRANSMISSION_SEQ_STEP", "TRANSMISSION_SEQ_STEP_ID")

	SQL = ""
	SQl = SQL & "SELECT * FROM TRANSMISSION_SEQ_STEP WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	set rs = conn.Execute(SQL)
	
	SQL2 = ""
	SQL2 = SQL2 & "SELECT * FROM OUTPUT_ITEM WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	set rs2 = conn.Execute(SQL2)
	
	SQLCopy = ""
	SQLCopy = SQLCopy & "INSERT INTO TRANSMISSION_SEQ_STEP ( TRANSMISSION_SEQ_STEP_ID, ROUTING_PLAN_ID, "
	SQLCopy = SQLCopy & "SEQUENCE, RETRY_COUNT, RETRY_WAIT_TIME, DESTINATION_STRING, TRANSMISSION_TYPE_ID, BATCH_HOLD, ALT_DESTINATION_STRING ) VALUES ("
	SQLCopy = SQLCopy & NextTrans & ", "
	SQLCopy = SQLCopy & RS("ROUTING_PLAN_ID") & ", "
	SQLCopy = SQLCopy & RS("SEQUENCE") & ", "
	SQLCopy = SQLCopy & SwapNull(RS("RETRY_COUNT")) & ", "
	SQLCopy = SQLCopy & SwapNull(RS("RETRY_WAIT_TIME")) & ", "
	SQLCopy = SQLCopy & "'" & RS("DESTINATION_STRING") & "', "
	SQLCopy = SQLCopy & RS("TRANSMISSION_TYPE_ID") & ","
	SQLCopy = SQLCopy & "'" & RS("BATCH_HOLD") & "', "
	SQLCopy = SQLCopy & "'" & RS("ALT_DESTINATION_STRING") & "') "

	set rs3 = conn.Execute(SQLCopy)
	
								   

	'------------------------JBOR-0065-------------------------------'
	 
	  
	 if  TRIM(RS("TRANSMISSION_TYPE_ID")) = "5" then
		
		nEDIOutboundItemID = NextPKey("EDI_OUTBOUND_ITEM", "EDI_OUTBOUND_ITEM_ID")
		nChar = chr(13) & chr(10)
	
		SQL = ""
		SQl = SQL & "SELECT * FROM EDI_OUTBOUND_ITEM WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
		set rsEDI = conn.Execute(SQL)
	   
		 Conn.BeginTrans
		 
			cSP1 = "INSERT INTO EDI_OUTBOUND_ITEM (EDI_OUTBOUND_ITEM_ID,TRANSMISSION_SEQ_STEP_ID,NAME,DESCRIPTION,DELIMITER,RECORD_DELIMITER) VALUES ('"& nEDIOutboundItemID &"','"& NextTrans &"','"& rsEDI("NAME") &"','',',','"& nChar &"')"
			Conn.Execute (cSP1)

			cSP2 = "INSERT INTO EDI_OUTBOUND_TOP_SEGMENTS( EDI_OUTBOUND_ITEM_ID,EDI_OUTBOUND_SEGMENT_ID) VALUES ('"& nEDIOutboundItemID &"','1')"
			Conn.Execute (cSP2)

		 Conn.CommitTrans  
		 
	end if
   '------------------------JBOR-0065-------------------------------'	

	SQL2 = ""
	SQL2 = SQL2 & "SELECT * FROM OUTPUT_ITEM WHERE TRANSMISSION_SEQ_STEP_ID=" & Request.QueryString("TRANSMISSION_SEQ_STEP_ID")
	set rs2 = conn.Execute(SQL2)
	
	Do WHile Not rs2.eof
	SQLCopy2 = ""
	SQLCopy2 = SQLCopy2 & "INSERT INTO OUTPUT_ITEM (OUTPUT_ITEM_ID, TRANSMISSION_SEQ_STEP_ID, OUTPUTDEF_ID, SEQUENCE) VALUES ("
	SQLCopy2 = SQLCopy2 & NextPKey("OUTPUT_ITEM", "OUTPUT_ITEM_ID") & ", "
	SQLCopy2 = SQLCopy2 & NextTrans & ", " 
	SQLCopy2 = SQLCopy2 & RS2("OUTPUTDEF_ID") & ", "
	SQLCopy2 = SQLCopy2 & RS2("SEQUENCE") & ")"
	set rs4 = conn.Execute(SQLCopy2)
	rs2.movenext
	loop
End If
End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
<% If Request.QueryString("STATUS") = "DELETE" OR Request.QueryString("STATUS") = "DELETE_OUTPUT_ITEM" OR Request.QueryString("STATUS") = "COPY" Then %>
	Parent.frames("WORKAREA").location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
<% Else %>
	parent.frames("TOP").document.all.StatusSpan.Style.Color = "#006699"
	parent.frames("TOP").document.all.StatusSpan.InnerHtml = "Transmission Saved"
<% End If %>
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>
