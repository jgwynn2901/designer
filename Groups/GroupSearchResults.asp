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
<title>Group Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.GID.value = ""
	end if
End Sub

Function GetGID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetGID = document.all.tblFields.rows(idx).getAttribute("GID")
	Else
		GetGID = ""
	End If
End Function

Function GetGIDName
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("NAME").innerText
	End If
	GetGIDName = strText
End Function
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Group Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
	dim RecCount
	RecCount = -1
	If Request.QueryString <> "" Then
	RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				GID = Request.QueryString("SearchGID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
			Case "C"
				GID = "%" & Request.QueryString("SearchGID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
			Case "E"
				GID = Request.QueryString("SearchGID")
				NAME = Request.QueryString("SearchName")
		End Select
	
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(GROUP_NM) LIKE '" & UCASE(NAME)  & "'"
		else
			WHERECLS = WHERECLS & "GROUP_NM LIKE '%'"
		End If
		If Request.QueryString("SearchGID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "GROUP_ID LIKE '" & GID & "'"
		End If
		
		if WHERECLS <> "" then
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT * FROM GROUPS WHERE "  & WHERECLS & " ORDER BY GROUP_NM" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" GID="">
	<td COLSPAN="3" NOWRAP CLASS="ResultCell">No groups found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" GID="<%=RS("GROUP_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("GROUP_ID"))%></td>
<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("GROUP_NM"))%></td>
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
End If
%>

</tbody>
</table>
</div>
</fieldset>
<script LANGUAGE="VBScript">
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
</script>
</body>
</html>
