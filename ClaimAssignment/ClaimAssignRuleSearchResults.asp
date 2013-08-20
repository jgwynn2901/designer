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
<title>Claim Number Assignment Rule Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.CARID.value = ""
	end if
End Sub

Function GetCARID
	GetCARID = getmultipleindex(document.all.tblFields, "CARID")
End Function
</script>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "CARID")
		return objRow.getAttribute("CARID");
		
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Claim Num Asgmt. Rule Id</div></td>
			<td class="thd"><div id><nobr>Seq</div></td>
			<td class="thd"><div id><nobr>A.H. Step ID</div></td>
			<td class="thd"><div id><nobr>Branch ID</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Rule Text</div></td>
			<td class="thd"><div id><nobr>Rule ID</div></td>
			<td class="thd"><div id><nobr>Attr Name</div></td>
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
				CARID = Request.QueryString("SearchCARID") & "%"
				BID = Request.QueryString("SearchBID") & "%"
				LOB = Request.QueryString("SearchLOBCD") & "%"
				AHSID = Request.QueryString("SearchAHSID") 
				RULEID = Request.QueryString("SearchRuleID") & "%"
				RULETEXT = Request.QueryString("SearchRuleText") & "%"
				ATTRNAME = Request.QueryString("SearchAttrName") & "%"
			Case "C"
				CARID = "%" & Request.QueryString("SearchCARID") & "%"
				BID = "%" & Request.QueryString("SearchBID") & "%"
				LOB = "%" & Request.QueryString("SearchLOBCD") & "%"
				AHSID =  Request.QueryString("SearchAHSID") 
				RULEID = "%" & Request.QueryString("SearchRuleID") & "%"
				RULETEXT = "%" & Request.QueryString("SearchRuleText") & "%"
				ATTRNAME = "%" & Request.QueryString("SearchAttrName") & "%"
			Case "E"
				CARID = Request.QueryString("SearchCARID")
				BID = Request.QueryString("SearchBID")
				LOB = Request.QueryString("SearchLOBCD")
				AHSID = Request.QueryString("SearchAHSID")
				RULEID = Request.QueryString("SearchRuleID")
				RULETEXT = Request.QueryString("SearchRuleText")
				ATTRNAME = Request.QueryString("SearchAttrName")
		End Select
	
		If Request.QueryString("SearchCARID") <> "" Then
			WHERECLS = WHERECLS & "CLAIMNUMBERASSIGNMENTRULE_ID LIKE '" & CARID & "'"
		End If
		If Request.QueryString("SearchBID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "BRANCH_ID LIKE '" & BID & "'"
		End If
		If Request.QueryString("SearchLOBCD") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "LOB_CD LIKE '" & LOB & "'"
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
			WHERECLS = WHERECLS & "CLAIMNUMBERASSIGNMENTRULE.RULE_ID LIKE '" & RULEID & "'"
		End If
		If Request.QueryString("SearchRuleText") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULE_TEXT) LIKE '" & UCASE(RULETEXT) & "'"
		End If
		
		If Request.QueryString("SearchAttrName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ASSIGN_TO_ATTR_NAME) LIKE '" & UCASE(ATTRNAME) & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT * FROM CLAIMNUMBERASSIGNMENTRULE, RULES WHERE " &_
				   "CLAIMNUMBERASSIGNMENTRULE.RULE_ID = RULES.RULE_ID(+) "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & "AND " & WHERECLS 
			End If
			
			SQLST = SQLST & " ORDER BY SEQUENCE, CLAIMNUMBERASSIGNMENTRULE_ID" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  CARID=''>
	<td COLSPAN="6"  NOWRAP CLASS="ResultCell" >No claim number assignment rules found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" CARID="<%=RS("CLAIMNUMBERASSIGNMENTRULE_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CLAIMNUMBERASSIGNMENTRULE_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SEQUENCE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("BRANCH_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LOB_CD"))%></td>
<td  TITLE="<%=ReplaceQuotesInText(renderCell(RS("RULE_TEXT")))%>" NOWRAP CLASS="ResultCell"><%=TruncateText(renderCell(RS("RULE_TEXT")),25)%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("RULE_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ASSIGN_TO_ATTR_NAME"))%></td>
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
