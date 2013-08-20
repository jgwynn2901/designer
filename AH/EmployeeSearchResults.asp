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
<title>Employee Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		'Parent.frames("TOP").document.all.EID.value = ""
	end if
End Sub

Function GetEID
	GetEID = getmultipleindex(document.all.tblFields, "EID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "EID")
		return objRow.getAttribute("EID");
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
			<td class="thd"><div id><nobr>Employee Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>SSN</div></td>
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
				EID = Request.QueryString("SearchEID") & "%"
				SSN = Request.QueryString("SearchSSN") & "%"
				NAME_LAST = Request.QueryString("SearchNAME_LAST") & "%"
				NAME_FIRST = Request.QueryString("SearchNAME_FIRST") & "%"
			Case "C"
				EID = "%" & Request.QueryString("SearchEID") & "%"
				SSN = "%" & Request.QueryString("SearchSSN") & "%"
				NAME_LAST = "%" & Request.QueryString("SearchNAME_LAST") & "%"
				NAME_FIRST = "%" & Request.QueryString("SearchNAME_FIRST") & "%"
			Case "E"
				EID = Request.QueryString("SearchEID")
				SSN = Request.QueryString("SearchSSN")
				NAME_LAST = Request.QueryString("SearchNAME_LAST")
				NAME_FIRST = Request.QueryString("SearchNAME_FIRST")
			End Select
		EID = Replace(EID, "'", "''")
		SSN = Replace(SSN, "'", "''")
		NAME_LAST = Replace(NAME_LAST, "'", "''")
		NAME_FIRST = Replace(NAME_FIRST, "'", "''")
		
		If Request.QueryString("SearchEID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(EMPLOYEE_ID) LIKE '" & UCASE(EID)  & "'"
		End If
		If Request.QueryString("SearchSSN") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "SSN LIKE '" & SSN & "'"
		End If
		If Request.QueryString("SearchNAME_LAST") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME_LAST) LIKE '" & UCASE(NAME_LAST) & "'"
		End If
		If Request.QueryString("SearchNAME_FIRST") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME_FIRST) LIKE '" & UCASE(NAME_FIRST) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT EMPLOYEE_ID,NAME_LAST,NAME_FIRST,SSN FROM EMPLOYEE "
			
			if WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS
			End if
			
			SQLST = SQLST & " ORDER BY NAME_LAST" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  AID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No Employees found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  EID='<%=RS("EMPLOYEE_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("EMPLOYEE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME_LAST"))%>, <%= RS("NAME_FIRST") %></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SSN"))%></td>
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

