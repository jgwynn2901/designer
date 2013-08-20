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
<title>Claim Key Office Code Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.KOCID.value = ""
	end if
End Sub

Function GetKOCID
	GetKOCID = getmultipleindex(document.all.tblFields, "KOCID")
End Function
</script>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "KOCID")
		return objRow.getAttribute("KOCID");
		
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Claim KOC Type ID</div></td>
			<td class="thd"><div id><nobr>KOC</div></td>
			<td class="thd"><div id><nobr>Sub KOC</div></td>
			<td class="thd"><div id><nobr>Active</div></td>
			<td class="thd"><div id><nobr>Sequence</div></td>
			<td class="thd"><div id><nobr>Next</div></td>
			<td class="thd"><div id><nobr>Minimum</div></td>
			<td class="thd"><div id><nobr>Maximum</div></td>
			<td class="thd"><div id><nobr>Notify Every</div></td>
			<td class="thd"><div id><nobr>Notified</div></td>
			<td class="thd"><div id><nobr>RPID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
	dim RecCount
	RecCount = -1
	If Request.QueryString("SEARCHTYPE") <> "" Then
		RecCount = 0
			
		KOCID = Request.QueryString("SearchKOCID")
		BID = Request.QueryString("SearchBID")
		AHSID = Request.QueryString("SearchAHSID") 
		BNUM = Request.QueryString("SearchBNUM") & "%"
		KOC = Request.QueryString("SearchKOC") & "%"
		
		SQLST = "SELECT * FROM CLAIM_KOC_ASSIGNMENT CKA "
	
		If Request.QueryString("SearchKOCID") <> "" Then
			WHERECLS = WHERECLS & "CLAIM_KOC_ID LIKE '" & KOCID & "'"
		End If
		If Request.QueryString("SearchBID") <> "" Then
			SQLST = "SELECT CKA.* FROM CLAIM_KOC_ASSIGNMENT CKA, BRANCH B "
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "B.BRANCH_NUMBER = CKA.BRANCH_NUMBER AND "
			WHERECLS = WHERECLS & "B.BRANCH_ID = " & BID 
		End If
		If Request.QueryString("SearchKOC") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "CKA.CLAIM_KOC LIKE '" & KOC & "'"
		End If
		If Request.QueryString("SearchAHSID") <> "" Then
			SQLST = "SELECT DISTINCT CKA.* FROM CLAIM_KOC_ASSIGNMENT CKA, BRANCH B "
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "B.BRANCH_NUMBER = CKA.BRANCH_NUMBER AND "
			WHERECLS = WHERECLS & "B.ACCOUNT_HIERARCHY_LOAD_ID = " & AHSID
		End If
		If Request.QueryString("SearchBNUM") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "CKA.BRANCH_NUMBER LIKE '" & BNUM & "'"
		End If
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		If WHERECLS <> "" Then
			SQLST = SQLST & " WHERE " & WHERECLS 
		End If
		
		SQLST = SQLST & " ORDER BY CKA.BRANCH_NUMBER, SEQ" 

		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText

		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  CARID=''>
	<td COLSPAN="6"  NOWRAP CLASS="ResultCell" >No claim key office codes found, re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" KOCID="<%=RS("CLAIM_KOC_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CLAIM_KOC_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("BRANCH_NUMBER"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CLAIM_KOC"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACTIVE_FLG"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SEQ"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NEXT"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MINIMUM"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MAXIMUM"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NOTIFY_EVERY"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NOTIFIED_NUM"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ROUTING_PLAN_ID"))%></td>
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
