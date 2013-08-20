<%
'***************************************************************
'displays table with form to enter/edit Mailbox Assignment Types
'
'$History: MailboxAssignTypeDetailsData.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:47p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'***************************************************************
%>
<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next

	BranchTextLen = 25
	RuleTextLen = 25
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Mailbox Assignment Type Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
function f_LastBARuleRecord
	if document.all.tblFields.Rows.Length <= 2 Then
		f_LastBARuleRecord = true
	else
		f_LastBARuleRecord = false
	end if
end Function
Function GetSelectedMARID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedMARID = document.all.tblFields.rows(idx).getAttribute("MARID")
	Else
		GetSelectedMARID = ""
	End If
End Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" rules=all  cellSpacing="0" scrolling="auto" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Sequence</div></td>
			<td class="thd"><div id><nobr>State</div></td>			
			<td class="thd"><div id><nobr>Rule Text</div></td>			
			<td class="thd"><div id><nobr>Mailbox #</div></td>
			<td class="thd"><div id><nobr>Mailbox ID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	MATID = CStr(Request.QueryString("MATID"))
	If MATID <> "NEW" And MATID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT MA.MAILBOX_ASSIGNMENT_RULE_ID,MA.LOB_CD,MA.SEQUENCE_NUM,R.RULE_TEXT,MA.ROUTING_STATE,M.MAILBOX_NUMBER, MA.MAILBOX_ID FROM " &_
				"MAILBOX_ASSIGNMENT_RULE MA, RULES R, MAILBOX M WHERE MA.RULE_ID = R.RULE_ID(+) AND " &_
				"MA.MAILBOX_ID = M.MAILBOX_ID(+) AND " &_				
				"MA.MAILBOX_ASSIGNMENT_TYPE_ID = " & MATID & " ORDER BY MA.SEQUENCE_NUM, MA.ROUTING_STATE" 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);" MARID="<%=RS("MAILBOX_ASSIGNMENT_RULE_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("LOB_CD"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("SEQUENCE_NUM"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ROUTING_STATE"))%></td>
	<td NOWRAP CLASS="ResultCell" TITLE="<%=ReplaceQuotesInText(RS("RULE_TEXT"))%>"><%=renderCell(TruncateText(RS("RULE_TEXT"),RuleTextLen))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("MAILBOX_NUMBER"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("MAILBOX_ID"))%></td>
	</tr>

<%
		RS.MoveNext
		Loop
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>

</tbody>
</table>
</div>
</BODY>
</HTML>


