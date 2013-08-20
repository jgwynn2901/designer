<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Lookup Type Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.LUTID.value = ""
	end if
End Sub

Function GetLUTID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetLUTID = document.all.tblFields.rows(idx).getAttribute("LUTID")
	Else
		GetLUTID = ""
	End If
End Function

Function GetLUTIDName
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("NAME").innerText
	End If
	GetLUTIDName = strText
End Function
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Lookup Type Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
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
				LUTID = Request.QueryString("SearchLUTID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
			Case "C"
				LUTID = "%" & Request.QueryString("SearchLUTID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
			Case "E"
				LUTID = Request.QueryString("SearchLUTID")
				NAME = Request.QueryString("SearchName")
		End Select
	
		LUTID = Replace(LUTID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchLUTID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "LU_TYPE_ID LIKE '" & LUTID & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT * FROM LU_TYPE "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS 
			End If
			SQLST = SQLST & " ORDER BY NAME" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" LUTID=''>
	<td COLSPAN="3" NOWRAP CLASS="ResultCell" >No lookup types found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" LUTID="<%=RS("LU_TYPE_ID")%>">
<td NOWRAP CLASS="ResultCell" ID="LU_TYPE_ID"><%=renderCell(RS("LU_TYPE_ID"))%></td>
<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
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
