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
<title>Agent Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.AID.value = ""
	end if
End Sub

Function GetAID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetAID = document.all.tblFields.rows(idx).getAttribute("AID")
	Else
		GetAID = ""
	End If
End Function

Function GetAIDName
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("NAME").innerText
	End If
	GetAIDName = strText
End Function
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Agent Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Agent Number</div></td>
			<td class="thd"><div id><nobr>Branch ID</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Zip</div></td>
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
				AID = Request.QueryString("SearchAID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				AGENTNO = Request.QueryString("AgentNo") & "%"
				STATE = Request.QueryString("SearchState") & "%"
				ZIP = Request.QueryString("SearchZip") & "%"
			Case "C"
				AID = "%" & Request.QueryString("SearchAID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
				AGENTNO = "%" & Request.QueryString("AgentNo") & "%"
				STATE = "%" & Request.QueryString("SearchState") & "%"
				ZIP = "%" & Request.QueryString("SearchZip") & "%"
			Case "E"
				AID = Request.QueryString("SearchAID")
				NAME = Request.QueryString("SearchName")
				AGENTNO = Request.QueryString("AgentNo")
				STATE = Request.QueryString("SearchState")
				ZIP = Request.QueryString("SearchZip")
		End Select
		AID = Replace(AID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		AGENTNO = Replace(AGENTNO, "'", "''")
		STATE = Replace(STATE, "'", "''")
		ZIP = Replace(ZIP, "'", "''")
		If Request.QueryString("SearchAID") <> "" Then
			WHERECLS = WHERECLS & "A.Agent_ID LIKE '" & AID & "'"
		End If
		If Request.QueryString("SearchName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(A.NAME) LIKE '" & UCase(NAME) & "'"
		End If
		If Request.QueryString("AgentNo") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(A.AGENT_NUMBER) LIKE '" & UCase(AGENTNO) & "'"
		End If
		If Request.QueryString("SearchState") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "A.STATE LIKE '" & STATE & "'"
		End If
		If Request.QueryString("SearchZip") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "A.ZIPCODE LIKE '" & ZIP & "'"
		End If		

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT A.AGENT_ID, A.NAME, A.AGENT_NUMBER, O.BRANCH_ID, " &_
					"A.STATE, A.ZIPCODE " &_
					"FROM Agent A, Branch O WHERE " &_
					"A.BRANCH_ID = O.BRANCH_ID(+) "
			If WHERECLS <> "" Then
				SQLST = SQLST & " AND " & WHERECLS
			End If
			SQLST = SQLST & " ORDER BY A.Agent_ID" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" AID=''>
	<td COLSPAN="6"  NOWRAP CLASS="ResultCell" >No agents found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" AID="<%=RS("Agent_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("Agent_ID"))%></td>
<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("AGENT_NUMBER"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("BRANCH_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ZIPCODE"))%></td>
</tr>
<%
				RS.MoveNext
				Loop
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing
End If
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
