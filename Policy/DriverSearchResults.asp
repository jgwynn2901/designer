<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"
%>
<!--#include file="..\lib\tablecommon.inc"-->
<% Response.Expires = 0 %>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Driver Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.DID.value = ""
	end if
End Sub

Function GetDID
	GetDID = getmultipleindex(document.all.tblFields, "DID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "DID")
		return objRow.getAttribute("DID");
	
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Driver Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>SSN</div></td>
			<td class="thd"><div id><nobr>License Number</div></td>
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
				DID = Request.QueryString("SearchDID") & "%"
				POLICY_ID = Request.QueryString("SearchPOLICY_ID") & "%"
				NAME_FIRST = Request.QueryString("SearchNAME_FIRST") & "%"
				NAME_LAST = Request.QueryString("SearchNAME_LAST") & "%"
				SSN = Request.QueryString("SearchSSN") & "%"
				ADDRESS = Request.QueryString("SearchADDRESS") & "%"
				CITY = Request.QueryString("SearchCITY") & "%"
				STATE = Request.QueryString("SearchSTATE") & "%"
				ZIP = Request.QueryString("SearchZIP") & "%"
			Case "C"
				DID = "%" & Request.QueryString("SearchDID") & "%"
				POLICY_ID = "%" & Request.QueryString("SearchPOLICY_ID") & "%"
				NAME_FIRST = "%" & Request.QueryString("SearchNAME_FIRST") & "%"
				NAME_LAST = "%" & Request.QueryString("SearchNAME_LAST") & "%"
				SSN = "%" & Request.QueryString("SearchSSN") & "%"
				ADDRESS = "%" & Request.QueryString("SearchADDRESS") & "%"
				CITY = "%" & Request.QueryString("SearchCITY") & "%"
				STATE = "%" & Request.QueryString("SearchSTATE") & "%"
				ZIP = "%" & Request.QueryString("SearchZIP") & "%"
			Case "E"
				DID = Request.QueryString("SearchDID")
				POLICY_ID = Request.QueryString("SearchPOLICY_ID")
				NAME_FIRST = Request.QueryString("SearchNAME_FIRST")
				NAME_LAST = Request.QueryString("SearchNAME_LAST")
				SSN = Request.QueryString("SearchSSN")
				ADDRESS = Request.QueryString("SearchADDRESS")
				CITY = Request.QueryString("SearchCITY")
				STATE = Request.QueryString("SearchSTATE")
				ZIP = Request.QueryString("SearchZIP")
		End Select
		DID = Replace(DID, "'", "''")
		POLICY_ID = Replace(POLICY_ID, "'", "''")
		NAME_FIRST = Replace(NAME_FIRST, "'", "''")
		NAME_LAST = Replace(NAME_LAST, "'", "''")
		SSN = Replace(SSN, "'", "''")
		ADDRESS = Replace(ADDRESS, "'", "''")
		CITY = Replace(CITY, "'", "''")
		STATE = Replace(STATE, "'", "''")
		ZIP = Replace(ZIP, "'", "''")
		If Request.QueryString("SearchPOLICY_ID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(POLICY_ID) LIKE '" & UCASE(POLICY_ID)  & "'"
		End If
		If Request.QueryString("SearchDID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "DRIVER_ID LIKE '" & DID & "'"
		End If
		If Request.QueryString("SearchNAME_FIRST") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME_FIRST) LIKE '" & UCASE(NAME_FIRST) & "'"
		End If
		If Request.QueryString("SearchNAME_LAST") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME_LAST) LIKE '" & UCASE(NAME_LAST) & "'"
		End If
		If Request.QueryString("SearchADDRESS") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ADDRESS1) LIKE '" & UCASE(ADDRESS) & "'"
		End If
		If Request.QueryString("SearchCITY") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CITY) LIKE '" & UCASE(CITY) & "'"
		End If
		If Request.QueryString("SearchSTATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(STATE) LIKE '" & UCASE(STATE) & "'"
		End If
		If Request.QueryString("SearchZIP") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ZIP) LIKE '" & UCASE(ZIP) & "'"
		End If
		If Request.QueryString("SearchSSN") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(SSN) LIKE '" & UCASE(SSN) & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT DRIVER_ID,NAME_FIRST,NAME_LAST, LICENSE_NUMBER, SSN FROM DRIVER "
			If WHERECLS <> "" Then
			SQLST = SQLST & "WHERE " & WHERECLS
			End if
			SQLST = SQLST & " ORDER BY NAME_LAST" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" DID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No drivers found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  DID='<%=RS("DRIVER_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DRIVER_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME_LAST"))%>, <%=RS("NAME_FIRST")%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SSN"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LICENSE_NUMBER"))%></td>
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

