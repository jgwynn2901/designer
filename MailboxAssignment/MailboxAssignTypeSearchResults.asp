<%
'***************************************************************
'display the results of a Mailbox Assignment Type search in table format
'
'$History: MailboxAssignTypeSearchResults.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:47p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'***************************************************************
%>
<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Mailbox Type Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.MATID.value = ""
	end if
End Sub

Function GetMATID
	GetMATID = getmultipleindex(document.all.tblFields, "MATID")
End Function
</script>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "MATID")
		return objRow.getAttribute("MATID");
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'95%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Mailbox Type Id</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>A.H. Step ID</div></td>
			<td class="thd"><div id><nobr>Rule Text</div></td>
			<td class="thd"><div id><nobr>Rule ID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%

	dim RecCount
	RecCount = -1
	If Request.QueryString("SEARCHTYPE") <> "" Then
	RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				MATID = Request.QueryString("SearchMATID") & "%"
				DESCRIPTION = Request.QueryString("SearchDescription") & "%"
				AHSID = Request.QueryString("SearchAHSID") 
				RULEID = Request.QueryString("SearchRuleID") & "%"
				RULETEXT = Request.QueryString("SearchRuleText") & "%"
			Case "C"
				MATID = "%" & Request.QueryString("SearchMATID") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDescription") & "%"
				AHSID =  Request.QueryString("SearchAHSID") 
				RULEID = "%" & Request.QueryString("SearchRuleID") & "%"
				RULETEXT = "%" & Request.QueryString("SearchRuleText") & "%"
			Case "E"
				MATID = Request.QueryString("SearchMATID")
				DESCRIPTION = Request.QueryString("SearchDescription")
				AHSID = Request.QueryString("SearchAHSID")
				RULEID = Request.QueryString("SearchRuleID")
				RULETEXT = Request.QueryString("SearchRuleText")
		End Select
		MATID = Replace(MATID, "'", "''")
		DESCRIPTION = Replace(DESCRIPTION, "'", "''")
		AHSID = Replace(AHSID, "'", "''")
		RULEID = Replace(RULEID, "'", "''")
		RULETEXT = Replace(RULETEXT, "'", "''")
		
		If Request.QueryString("SearchMATID") <> "" Then
			WHERECLS = WHERECLS & "MAILBOX_ASSIGNMENT_TYPE_ID LIKE '" & MATID & "'"
		End If
		If Request.QueryString("SearchDescription") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		If Request.QueryString("SearchAHSID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID = "  &  AHSID 
		End If
		If Request.QueryString("SearchRuleID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "MAILBOX_ASSIGNMENT_TYPE.RULE_ID LIKE '" & RULEID & "'"
		End If
		If Request.QueryString("SearchRuleText") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULE_TEXT) LIKE '" & UCASE(RULETEXT) & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT * FROM MAILBOX_ASSIGNMENT_TYPE, RULES WHERE " &_
				   "MAILBOX_ASSIGNMENT_TYPE.RULE_ID = RULES.RULE_ID(+) " 
				   
			If WHERECLS <> "" Then
				SQLST = SQLST & " AND " & WHERECLS 
			End If
			
			SQLST = SQLST & " ORDER BY MAILBOX_ASSIGNMENT_TYPE_ID" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" MATID=''>
	<td COLSPAN="5" NOWRAP CLASS="ResultCell">No Mailbox assignment types found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" MATID="<%=RS("MAILBOX_ASSIGNMENT_TYPE_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MAILBOX_ASSIGNMENT_TYPE_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DESCRIPTION"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
<td  TITLE="<%=ReplaceQuotesInText(renderCell(RS("RULE_TEXT")))%>" NOWRAP CLASS="ResultCell"><%=TruncateText(renderCell(RS("RULE_TEXT")),25)%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("RULE_ID"))%></td>
</tr>
<%
				RS.MoveNext
				Loop
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing

		End if

%>

</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">
<%	If RecCount >= 0 Then %>
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
<%	End If %>
</SCRIPT>
</body>
</html>
