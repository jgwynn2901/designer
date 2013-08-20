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
<title>Managed Care Branch Assignment Type Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.MCBATID.value = ""
	end if
End Sub

Function GetMCBATID
	GetMCBATID = getmultipleindex(document.all.tblFields, "MCBATID")
'	dim idx
'	idx = CInt(getselectedindex(document.all.tblFields))
'	If idx <> -1 Then
'		GetBATID = document.all.tblFields.rows(idx).getAttribute("BATID")
'	Else
'		GetBATID = ""
'	End If
End Function
</script>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "MCBATID")
		return objRow.getAttribute("MCBATID");
		
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>MC Branch Asgmt. Type Id</div></td>
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
				MCBATID = Request.QueryString("SearchMCBATID") & "%"
				DESCRIPTION = Request.QueryString("SearchDescription") & "%"
				AHSID = Request.QueryString("SearchAHSID") 
				RULEID = Request.QueryString("SearchRuleID") & "%"
				RULETEXT = Request.QueryString("SearchRuleText") & "%"
			Case "C"
				MCBATID = "%" & Request.QueryString("SearchMCBATID") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDescription") & "%"
				AHSID =  Request.QueryString("SearchAHSID") 
				RULEID = "%" & Request.QueryString("SearchRuleID") & "%"
				RULETEXT = "%" & Request.QueryString("SearchRuleText") & "%"
			Case "E"
				MCBATID = Request.QueryString("SearchMCBATID")
				DESCRIPTION = Request.QueryString("SearchDescription")
				AHSID = Request.QueryString("SearchAHSID")
				RULEID = Request.QueryString("SearchRuleID")
				RULETEXT = Request.QueryString("SearchRuleText")
		End Select
		MCBATID = Replace(MCBATID, "'", "''")
		DESCRIPTION = Replace(DESCRIPTION, "'", "''")
		AHSID = Replace(AHSID, "'", "''")
		RULEID = Replace(RULEID, "'", "''")
		RULETEXT = Replace(RULETEXT, "'", "''")
		
		If Request.QueryString("SearchMCBATID") <> "" Then
			WHERECLS = WHERECLS & "MC_BRANCHASSIGNMENTTYPE_ID LIKE '" & MCBATID & "'"
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
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID = " & AHSID 
		End If
		If Request.QueryString("SearchRuleID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "MC_BRANCHASSIGNMENTTYPE.RULE_ID LIKE '" & RULEID & "'"
		End If
		If Request.QueryString("SearchRuleText") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULE_TEXT) LIKE '" & UCASE(RULETEXT) & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT * FROM MC_BRANCHASSIGNMENTTYPE, RULES WHERE " &_
				   "MC_BRANCHASSIGNMENTTYPE.RULE_ID = RULES.RULE_ID(+) " 
				   
			If WHERECLS <> "" Then
				SQLST = SQLST & " AND " & WHERECLS 
			End If
			
			SQLST = SQLST & " ORDER BY MC_BRANCHASSIGNMENTTYPE_ID" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" MCBATID=''>
	<td COLSPAN="5" NOWRAP CLASS="ResultCell">No managed care branch assignment types found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" MCBATID="<%=RS("MC_BRANCHASSIGNMENTTYPE_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MC_BRANCHASSIGNMENTTYPE_ID"))%></td>
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
