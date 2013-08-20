<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"
%>

<!--#include file="..\lib\tablecommon.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Office Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.OID.value = ""
	end if
End Sub

Function GetOID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetOID = document.all.tblFields.rows(idx).getAttribute("OID")
	else 
		GetOID = ""
	End If
End Function

Function GetOIDNumber
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("NUMBER").innerText
	End If
	GetOIDNumber = strText
End Function

Function GetOIDState
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("STATE").innerText
	End If
	GetOIDState = strText
End Function

Function GetOIDType
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("TYPE").innerText
	End If
	GetOIDType = strText
End Function

Function GetOIDZip
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("ZIP").innerText
	End If
	GetOIDZip = strText
End Function

</script>
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin=0 leftmargin=0  rightmargin=0 >
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:100%;width:100%">
<div align="LEFT" style="{display:block;height:100%;width:100%;overflow:auto}">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Office Id</div></td>
			<td class="thd"><div id><nobr>Number</div></td>
			<td class="thd"><div id><nobr>Type</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Zip</div></td>
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
				OID = Request.QueryString("SearchOID") & "%"
				NUMBER = Request.QueryString("SearchNumber") & "%"
				STATE = Request.QueryString("SearchState") & "%"
				OTYPE = Request.QueryString("SearchOType") & "%"
				ZIP = Request.QueryString("SearchZip") & "%"
			Case "C"
				OID = "%" & Request.QueryString("SearchOID") & "%"
				NUMBER = "%" & Request.QueryString("SearchNumber") & "%"
				STATE = "%" & Request.QueryString("SearchState") & "%"
				OTYPE = "%" & Request.QueryString("SearchOType") & "%"
				ZIP = "%" & Request.QueryString("SearchZip") & "%"
			Case "E"
				OID = Request.QueryString("SearchOID")
				NUMBER = Request.QueryString("SearchNumber")
				STATE = Request.QueryString("SearchState")
				OTYPE = Request.QueryString("SearchOType")
				ZIP = Request.QueryString("SearchZip")
		End Select

	
		OID = Replace(OID,"'","''")
		NUMBER = Replace(NUMBER,"'","''")
		STATE = Replace(STATE,"'","''")
		OTYPE = Replace(OTYPE,"'","''")
		ZIP = Replace(ZIP,"'","''")
		
		If Request.QueryString("SearchOID") <> "" Then
			WHERECLS = WHERECLS & "PK_OFFICE LIKE '" & OID  & "'"
		End If
		If Request.QueryString("SearchNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "OFFICE_NUMBER LIKE '" & NUMBER  & "'"
		End If
		If Request.QueryString("SearchState") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(STATE) LIKE '" & UCASE(STATE) & "'"
		End If
		If Request.QueryString("SearchOType") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(OFFICE_TYPE) LIKE '" & UCASE(OTYPE) & "'"
		End If
		If Request.QueryString("SearchZip") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ZIP) LIKE '" & UCASE(ZIP) & "'"
		End If

		if WHERECLS <> "" then
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT * FROM Office WHERE " & WHERECLS & " ORDER BY PK_OFFICE" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText

			If RS.EOF And RS.BOF then
%>
<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  OID='' >
	<td COLSPAN=5  NOWRAP CLASS="ResultCell" >No offices found re-check your criteria</td>
</tr>
<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OID='<%=RS("PK_OFFICE")%>'  >
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("PK_OFFICE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NUMBER"><%=renderCell(RS("OFFICE_NUMBER"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="TYPE" ><%=renderCell(RS("OFFICE_TYPE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="STATE"><%=renderCell(RS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="ZIP" ><%=renderCell(RS("ZIP"))%></td>

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
