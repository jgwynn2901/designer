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
<title>Carrier Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.CID.value = ""
	end if
End Sub

Function GetCID
	GetCID = getmultipleindex(document.all.tblFields, "CID")
End Function

Function GetCIDName
	GetCIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function


</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "CID")
		return objRow.getAttribute("CID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
		
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd" width="4"><div id><nobr>Carrier ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd" width="4"><div id><nobr>Carrier Number</div></td>
			<td class="thd" width="2"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Title</div></td>
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
				CID = Request.QueryString("SearchCID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				CARRIERNUMBER = Request.QueryString("SearchCarrierNumber") & "%"
				TITLE = Request.QueryString("SearchTITLE") & "%"
				BUSINESSTYPE = Request.QueryString("SearchBUSINESSTYPE") & "%"
				ADDRESS = Request.QueryString("SearchADDRESS") & "%"
				CITY = Request.QueryString("SearchCITY") & "%"
				STATE = Request.QueryString("SearchSTATE") & "%"
				ZIP = Request.QueryString("SearchZIP") & "%"
			Case "C"
				CID = "%" & Request.QueryString("SearchCID") & "%"
				NAME = "%" & Request.QueryString("SearchNAME") & "%"
				CARRIERNUMBER = "%" & Request.QueryString("SearchCarrierNumber") & "%"
				TITLE = "%" & Request.QueryString("SearchTITLE") & "%"
				BUSINESSTYPE = "%" & Request.QueryString("SearchBUSINESSTYPE") & "%"
				ADDRESS = "%" & Request.QueryString("SearchADDRESS") & "%"
				CITY = "%" & Request.QueryString("SearchCITY") & "%"
				STATE = "%" & Request.QueryString("SearchSTATE") & "%"
				ZIP = "%" & Request.QueryString("SearchZIP") & "%"
			Case "E"
				CID = Request.QueryString("SearchAID")
				NAME = Request.QueryString("SearchName")
				CARRIERNUMBER = Request.QueryString("SearchCarrierNumber")
				TITLE = Request.QueryString("SearchTITLE")
				BUSINESSTYPE = Request.QueryString("SearchBUSINESSTYPE")
				ADDRESS = Request.QueryString("SearchADDRESS")
				CITY = Request.QueryString("SearchCITY")
				STATE = Request.QueryString("SearchSTATE")
				ZIP = Request.QueryString("SearchZIP")
		End Select
		CID = Replace(CID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		CARRIERNUMBER = Replace(CARRIERNUMBER, "'", "''")
		TITLE = Replace(TITLE, "'", "''")
		BUSINESSTYPE = Replace(BUSINESSTYPE, "'", "''")
		ADDRESS = Replace(ADDRESS, "'", "''")
		CITY = Replace(CITY, "'", "''")
		STATE = Replace(STATE, "'", "''")
		ZIP = Replace(ZIP, "'", "''")
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchCID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "CARRIER_ID LIKE '" & CID & "'"
		End If
		If Request.QueryString("SearchCarrierNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "CARRIER_NUMBER LIKE '" & CARRIERNUMBER & "'"
		End If		
		If Request.QueryString("SearchTITLE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(TITLE) LIKE '" & UCASE(TITLE) & "'"
		End If
		If Request.QueryString("SearchBUSINESSTYPE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(BUSINESS_TYPE) LIKE '" & UCASE(BUSINESS_TYPE) & "'"
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
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT CARRIER_NUMBER, NAME, TITLE, STATE, CARRIER_ID FROM CARRIER "
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS
			End If
			 SQLST = SQLST & " ORDER BY NAME" 
			
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  CID='' >
	<td COLSPAN=10 NOWRAP CLASS="ResultCell">No carriers found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  CID='<%=RS("CARRIER_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CARRIER_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CARRIERNUMBER"><%=renderCell(RS("CARRIER_NUMBER"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("TITLE"))%></td>
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
