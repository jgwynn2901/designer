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
<title>Vendor Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.VID.value = ""
	end if
End Sub

Function GetVID
	GetVID = getmultipleindex(document.all.tblFields, "VID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "VID")
		return objRow.getAttribute("VID");
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'90%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0"  rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Vendor ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Address</div></td>			
			<td class="thd"><div id><nobr>City</div></td>			
			<td class="thd"><div id><nobr>State</div></td>			
			<td class="thd"><div id><nobr>Zip</div></td>			
			<td class="thd"><div id><nobr>Enabled</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount
	RecCount = -1
	WHERECLS = ""
	If Request.QueryString <> "" Then
		RecCount = 0
		if UCase(Request.QueryString("SearchEnabled")) = "ON" then
			ENABLED = "Y"
		else
			ENABLED = "N"
		end if
		cVerb = "LIKE"
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				VID = Request.QueryString("SearchVendorID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				ADDRESS = Request.QueryString("SearchAddress") & "%"
				CITY = Request.QueryString("SearchCity") & "%"
				STATE = Request.QueryString("SearchState") & "%"
				ZIP = Request.QueryString("SearchZip") & "%"
				SERVICETYPE = Request.QueryString("SearchServType") & "%"
			Case "C"
				VID = "%" & Request.QueryString("SearchVendorID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
				ADDRESS = "%" & Request.QueryString("SearchAddress") & "%"
				CITY = "%" & Request.QueryString("SearchCity") & "%"
				STATE = "%" & Request.QueryString("SearchState") & "%"
				ZIP = "%" & Request.QueryString("SearchZip") & "%"
				SERVICETYPE = "%" & Request.QueryString("SearchServType") & "%"
			Case "E"
				cVerb = "="
				VID = Request.QueryString("SearchVendorID")
				NAME = Request.QueryString("SearchName")
				ADDRESS = Request.QueryString("SearchAddress")
				CITY = Request.QueryString("SearchCity")
				STATE = Request.QueryString("SearchState")
				ZIP = Request.QueryString("SearchZip")
				SERVICETYPE = Request.QueryString("SearchServType")
		End Select
		NAME = Replace(NAME, "'", "''")
		ADDRESS = Replace(ADDRESS, "'", "''")
		CITY = Replace(CITY, "'", "''")
		STATE = Replace(STATE, "'", "''")
	
		If Request.QueryString("SearchVendorID") <> "" Then
			WHERECLS = WHERECLS & "VENDOR_ID " & cVerb & "'" & VID  & "'"
		End If
		If Request.QueryString("SearchName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "NAME " & cVerb & "'" & NAME & "'"
		End If
		If Request.QueryString("SearchAddress") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ADDRESS_1) " & cVerb & "'" & UCASE(ADDRESS) & "'"
		End If		
		If Request.QueryString("SearchCity") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CITY) " & cVerb & "'" & UCASE(CITY) & "'"
		End If
		If Request.QueryString("SearchState") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "STATE " & cVerb & "'" & STATE & "'"
		End If
		If Request.QueryString("SearchZip") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ZIP " & cVerb & "'" & ZIP & "'"
		End If		
		If WHERECLS <> "" Then 
			WHERECLS = WHERECLS & " AND "
		End If
		WHERECLS = WHERECLS & "ENABLED_FLG = '" & ENABLED & "'"
	
		lWithServiceType = false
		If Request.QueryString("SearchServType") <> "" Then 
			lWithServiceType = true
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ST.SERVICE_TYPE_ID=" & Request.QueryString("SearchServType") & _
			" AND ST.SERVICE_TYPE_ID=VS.SERVICE_TYPE_ID" & _
			" AND VS.VENDOR_ID=VENDOR.VENDOR_ID"
		End If
			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM VENDOR "
		if lWithServiceType then
			SQLST = SQLST & ",SERVICE_TYPE ST,VENDOR_SERVICE VS"
		end if
		if WHERECLS <> "" then SQLST = SQLST & " WHERE " & WHERECLS
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  BID='' >
	<td COLSPAN=8 NOWRAP CLASS="ResultCell" >No Vendors found re-check your criteria</td>
</tr>
	
<%		Else
			Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  VID='<%=RS("VENDOR_ID")%>'>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("VENDOR_ID"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS_1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CITY"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATE"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ZIP"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ENABLED_FLG"))%></td>
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
