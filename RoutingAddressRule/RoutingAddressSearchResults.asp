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
<title>Routing Address Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.RAID.value = ""
	end if
End Sub

Function GetRAID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetRAID = document.all.tblFields.rows(idx).getAttribute("RAID")
	else 
		GetRAID = ""
	End If
End Function

Function GetRAIDDescription
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("DESCRIPTION").innerText
	End If
	GetRAIDDescription = strText
End Function

Function GetRAIDState
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("STATE").innerText
	End If
	GetRAIDState = strText
End Function

Function GetRAIDFIPS
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("FIPS").innerText
	End If
	GetRAIDFIPS = strText
End Function

Function GetRAIDZip
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("ZIP").innerText
	End If
	GetRAIDZip = strText
End Function

</script>
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin=0 leftmargin=0  rightmargin=0 >
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:100%;width:100%">
<div align="LEFT" style="{display:block;height:100%;width:100%;overflow:auto}">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Routing Address Id</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>FIPS</div></td>
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
				RAID = Request.QueryString("SearchRAID") & "%"
				DESCRIPTION = Request.QueryString("SearchDescription") & "%"
				STATE = Request.QueryString("SearchState") & "%"
				FIPS = Request.QueryString("SearchFIPS") & "%"
				ZIP = Request.QueryString("SearchZip") & "%"
			Case "C"
				RAID = "%" & Request.QueryString("SearchRAID") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDescription") & "%"
				STATE = "%" & Request.QueryString("SearchState") & "%"
				FIPS = "%" & Request.QueryString("SearchFIPS") & "%"
				ZIP = "%" & Request.QueryString("SearchZip") & "%"
			Case "E"
				RAID = Request.QueryString("SearchRAID")
				DESCRIPTION = Request.QueryString("SearchDescription")
				STATE = Request.QueryString("SearchState")
				FIPS = Request.QueryString("SearchFIPS")
				ZIP = Request.QueryString("SearchZip")
		End Select

	
		RAID = Replace(RAID,"'","''")
		DESCRIPTION = Replace(DESCRIPTION,"'","''")
		STATE = Replace(STATE,"'","''")
		FIPS = Replace(FIPS,"'","''")
		ZIP = Replace(ZIP,"'","''")
		
		If Request.QueryString("SearchRAID") <> "" Then
			WHERECLS = WHERECLS & "ROUTINGADDRESS_ID LIKE '" & RAID  & "'"
		End If
		If Request.QueryString("SearchDescription") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION)  & "'"
		End If
		If Request.QueryString("SearchState") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(STATE) LIKE '" & UCASE(STATE) & "'"
		End If
		If Request.QueryString("SearchFIPS") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(FIPS) LIKE '" & UCASE(FIPS) & "'"
		End If
		If Request.QueryString("SearchZip") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ZIP) LIKE '" & UCASE(ZIP) & "'"
		End If
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT * FROM ROUTINGADDRESS "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS 
			End if
			
			SQLST = SQLST & " ORDER BY ROUTINGADDRESS_ID" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText

			If RS.EOF And RS.BOF then
%>
<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  RAID='' >
	<td COLSPAN=5  NOWRAP CLASS="ResultCell" >No routing addresses found re-check your criteria</td>
</tr>
<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" RAID='<%= RS("ROUTINGADDRESS_ID") %>'  >
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ROUTINGADDRESS_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="DESCRIPTION"><%=renderCell(RS("DESCRIPTION"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="STATE"><%=renderCell(RS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="FIPS" ><%=renderCell(RS("FIPS"))%></td>
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
