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
		'Parent.frames("TOP").document.all.SearchODID.value = ""
	end if
End Sub

Function GetODID
	GetODID = getmultipleindex(document.all.tblFields, "ODID")
End Function

Function GetODIDName
	GetODIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "ODID")
		return objRow.getAttribute("ODID");
	if (whichCol == "NAME")
		return objRow.cells("NAME").innerText;
	
}
</SCRIPT>
</head>


<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="#d6cfbd">
<SPAN ID=SPANSTATUS></SPAN>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Output Def. ID</div></td>
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
				OUTPUTDEF_ID = Request.QueryString("SEARCHODID") & "%"
				NAME = Request.QueryString("SearchNAME") & "%"
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION") & "%"
			Case "C"
				OUTPUTDEF_ID = "%" & Request.QueryString("SEARCHODID") & "%"
				NAME = "%" & Request.QueryString("SearchNAME") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDESCRIPTION") & "%"
			Case "E"
				OUTPUTDEF_ID = Request.QueryString("SEARCHODID")
				NAME = Request.QueryString("SearchNAME")
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION")
			End Select
		OUTPUTDEF_ID = Replace(OUTPUTDEF_ID, "'", "")
		NAME =  Replace(NAME, "'", "''")
		DESCRIPTION =  Replace(DESCRIPTION, "'", "''")
	
		If Request.QueryString("SearchNAME") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SEARCHODID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "OUTPUTDEF_ID LIKE '" & OUTPUTDEF_ID & "'"
		End If
		If Request.QueryString("SEARCH_DESCRIPTION") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
	
		Set RS = Server.CreateObject("ADODB.RecordSet")
		RS.MaxRecords = MAXRECORDCOUNT
		ConnectionString = CONNECT_STRING
		SQLST = SQLST & "SELECT * FROM OUTPUT_DEFINITION "
		If WHERECLS <> "" Then
			SQLST = SQLST & "WHERE " & WHERECLS
		End If
		SQLST = SQLST & " ORDER BY NAME" 
		RS.Open SQLST, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" ODID='' >
	<td COLSPAN=3 NOWRAP CLASS="LABEL" ID="OUTPUTDEF_ID">No Output Definitions found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  ODID='<%=RS("OUTPUTDEF_ID")%>'>
	<td NOWRAP CLASS="ResultCell" ID="OUTPUTDEF_ID"><%=renderCell(RS("OUTPUTDEF_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="DESCRIPTION" ><%=renderCell(RS("DESCRIPTION"))%></td>
	</tr>

<%
				RS.MoveNext
				Loop
			End If
End If
%>

</tbody>
</table>
</div>
</fieldset>
<% If Request.QueryString <> "" Then %>
<SCRIPT LANGUAGE="VBScript">
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	elseif curCount = -1 then
		Parent.frames("TOP").UpdateStatus("<%=MSG_PROMPT%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
</SCRIPT>
<% End If %>
</body>
</html>
