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
<title>Address Book Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.ABID.value = ""
	end if
End Sub

Function GetABID
	GetABID = getmultipleindex(document.all.tblFields, "ABID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "ABID")
		return objRow.getAttribute("ABID");
	
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Address Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
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
				ABID = Request.QueryString("SearchABID") & "%"
				NAME = Request.QueryString("SearchNAME") & "%"
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION") & "%"
				CFID = Request.QueryString("SearchCFID") & "%"
			Case "C"
				ABID = "%" & Request.QueryString("SearchABID") & "%"
				NAME = "%" & Request.QueryString("SearchNAME") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDESCRIPTION") & "%"
				CFID = "%" & Request.QueryString("SearchCFID") & "%"
			Case "E"
				ABID = Request.QueryString("SearchABID")
				NAME = Request.QueryString("SearchNAME")
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION")
				CFID = Request.QueryString("SearchCFID")
		End Select
		ABID = Replace(ABID, "'", "''")	
		NAME = Replace(NAME, "'", "''")
		DESCRIPTION = Replace(DESCRIPTION, "'", "''")
		CFID = Replace(CFID, "'", "''")
		If Request.QueryString("SearchABID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(ADDRESS_BOOK_ENTRY_ID) LIKE '" & UCASE(ABID)  & "'"
		End If
		If Request.QueryString("SearchNAME") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME) & "'"
		End If
		If Request.QueryString("SearchDESCRIPTION") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		If Request.QueryString("SearchCFID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CALLFLOW_ID) LIKE '" & UCASE(CFID) & "'"
		End If
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT ADDRESS_BOOK_ENTRY_ID, NAME, DESCRIPTION FROM ADDRESS_BOOK_ENTRY "
			If WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS
			End If
			SQLST = SQLST & " ORDER BY NAME" 
			
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" ABID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No address book entries found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  ABID='<%=RS("ADDRESS_BOOK_ENTRY_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS_BOOK_ENTRY_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DESCRIPTION"))%></td>
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

