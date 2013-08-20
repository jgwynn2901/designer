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
<title>Output Overflow Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.OOID.value = ""
	end if
End Sub

Function GetOOID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetOOID = document.all.tblFields.rows(idx).getAttribute("OOID")
	Else
		GetOOID = ""
	End If
End Function
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Overflow Id</div></td>
			<td class="thd"><div id><nobr>A.H. Step ID</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Sequence</div></td>
			<td class="thd"><div id><nobr>Attribute Name</div></td>
			<td class="thd"><div id><nobr>Show when empty?</div></td>
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
				OOID = Request.QueryString("SearchOOID") & "%"
				LOBCD = Request.QueryString("SearchLOBCD") 
				AHSID = Request.QueryString("SearchAHSID") 
				NAME = Request.QueryString("SearchAttributeName") & "%"
				SEQUENCE = Request.QueryString("SearchSequence") & "%"
			Case "C"
				OOID = "%" & Request.QueryString("SearchOOID") & "%"
				LOBCD =  Request.QueryString("SearchLOBCD") 
				AHSID = Request.QueryString("SearchAHSID") 
				NAME = "%" & Request.QueryString("SearchAttributeName") & "%"
				SEQUENCE = "%" & Request.QueryString("SearchSequence") & "%"
			Case "E"
				OOID = Request.QueryString("SearchOOID")
				LOBCD = Request.QueryString("SearchLOBCD")
				AHSID = Request.QueryString("SearchAHSID")
				NAME = Request.QueryString("SearchAttributeName")
				SEQUENCE = Request.QueryString("SearchSequence")
		End Select
	
		OOID = Replace(OOID,"'","''")
		AHSID = Replace(AHSID,"'","''")
		NAME = Replace(NAME,"'","''")
		SEQUENCE = Replace(SEQUENCE,"'","''")
		
	
		If Request.QueryString("SearchOOID") <> "" Then
			WHERECLS = WHERECLS & "OUTPUT_OVERFLOW_ID LIKE '" & OOID & "'"
		End If
		If Request.QueryString("SearchLOBCD") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "LOB_CD = '" & LOBCD & "'"
		End If
		If Request.QueryString("SearchAHSID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID = " & AHSID 
		End If
		If Request.QueryString("SearchAttributeName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ATTRIBUTE_NAME) LIKE '" & UCase(NAME) & "'"
		End If
		If Request.QueryString("SearchSequence") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "SEQUENCE LIKE '" & SEQUENCE & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT * FROM OUTPUT_OVERFLOW "
			If WHERECLS <> "" Then
			SQLST = SQLST & "WHERE " & WHERECLS 
			End If
			SQLST = SQLST & " ORDER BY ACCNT_HRCY_STEP_ID, LOB_CD,SEQUENCE" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OOID=''>
	<td COLSPAN="6"  NOWRAP CLASS="ResultCell" >No output overflows found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OOID="<%=RS("OUTPUT_OVERFLOW_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("OUTPUT_OVERFLOW_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LOB_CD"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SEQUENCE"))%></td>
<td  TITLE="<%=ReplaceQuotesInText(renderCell(RS("ATTRIBUTE_NAME")))%>" NOWRAP CLASS="ResultCell"><%=TruncateText(renderCell(RS("ATTRIBUTE_NAME")),25)%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SHOW_WHEN_EMPTY_FLAG"))%></td>
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
