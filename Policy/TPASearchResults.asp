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
<title>TPA Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.TPAID.value = ""
	end if
End Sub

Function GetTPAID
	GetTPAID = getmultipleindex(document.all.tblFields, "TPAID")
End Function

Function GetTPAIDName
	GetTPAIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function


</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "TPAID")
		return objRow.getAttribute("TPAID");
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
			<td class="thd" width="4"><div id><nobr>TPA ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd" width="4"><div id><nobr>TPA Number</div></td>
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
				TPAID = Request.QueryString("SearchTPAID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				TPANUMBER = Request.QueryString("SearchTPANumber") & "%"
				TITLE = Request.QueryString("SearchTITLE") & "%"
				BUSINESSTYPE = Request.QueryString("SearchBUSINESSTYPE") & "%"
				ADDRESS = Request.QueryString("SearchADDRESS") & "%"
				CITY = Request.QueryString("SearchCITY") & "%"
				STATE = Request.QueryString("SearchSTATE") & "%"
				ZIP = Request.QueryString("SearchZIP") & "%"
			Case "C"
				TPAID = "%" & Request.QueryString("SearchTPAID") & "%"
				NAME = "%" & Request.QueryString("SearchNAME") & "%"
				TPANUMBER = "%" & Request.QueryString("SearchTPANumber") & "%"
				TITLE = "%" & Request.QueryString("SearchTITLE") & "%"
				BUSINESSTYPE = "%" & Request.QueryString("SearchBUSINESSTYPE") & "%"
				ADDRESS = "%" & Request.QueryString("SearchADDRESS") & "%"
				CITY = "%" & Request.QueryString("SearchCITY") & "%"
				STATE = "%" & Request.QueryString("SearchSTATE") & "%"
				ZIP = "%" & Request.QueryString("SearchZIP") & "%"
			Case "E"
				TPAID = Request.QueryString("SearchAID")
				NAME = Request.QueryString("SearchName")
				TPANUMBER = Request.QueryString("SearchTPANumber")
				TITLE = Request.QueryString("SearchTITLE")
				BUSINESSTYPE = Request.QueryString("SearchBUSINESSTYPE")
				ADDRESS = Request.QueryString("SearchADDRESS")
				CITY = Request.QueryString("SearchCITY")
				STATE = Request.QueryString("SearchSTATE")
				ZIP = Request.QueryString("SearchZIP")
		End Select
		TPAID = Replace(TPAID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		TPANUMBER = Replace(TPANUMBER, "'", "''")
		TITLE = Replace(TITLE, "'", "''")
		BUSINESSTYPE = Replace(BUSINESSTYPE, "'", "''")
		ADDRESS = Replace(ADDRESS, "'", "''")
		CITY = Replace(CITY, "'", "''")
		STATE = Replace(STATE, "'", "''")
		ZIP = Replace(ZIP, "'", "''")
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchTPAID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "TPA_ID LIKE '" & TPAID & "'"
		End If
		If Request.QueryString("SearchTPANumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "TPA_NUMBER LIKE '" & TPANUMBER & "'"
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
			SQLST =  "SELECT TPA_NUMBER, NAME, TITLE, STATE, TPA_ID FROM THIRD_PARTY_ADMINISTRATOR "
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS
			End If
			 SQLST = SQLST & " ORDER BY NAME" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenForwardOnly,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  TPAID='' >
	<td COLSPAN=10 NOWRAP CLASS="ResultCell">No TPAs found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  TPAID='<%=RS("TPA_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("TPA_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="TPANUMBER"><%=renderCell(RS("TPA_NUMBER"))%></td>
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
