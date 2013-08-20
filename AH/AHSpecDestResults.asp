<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<%
'	main code
'option explicit

dim oConn, cSQL, ConnectionString
dim cLastName, cFirstName, cMI, cSD, cPhone, cState, cCity, cZip
dim cWhere, oRS
dim cCN
dim lFullSearch

cCN = Request.QueryString("CLIENT_NODE")
lFullSearch = instr(1, Request.QueryString, "FULLSEARCH") <> 0
If Request.QueryString("SEARCHTYPE") <> "" Then
	Select Case Request.QueryString("SEARCHTYPE")
		Case "B"
			cLastName = Request.QueryString("LNAME") & "%"
			cFirstName = Request.QueryString("FNAME") & "%"
			cMI = Request.QueryString("MI") & "%"
			cSD = Request.QueryString("SD") & "%"
			cPhone = Request.QueryString("PHONE") & "%"
			cState = Request.QueryString("STATE") & "%"
			cCity = Request.QueryString("CITY") & "%"
			cZip = Request.QueryString("ZIP") & "%"
		Case "C"
			cLastName = "%" & Request.QueryString("LNAME") & "%"
			cFirstName = "%" & Request.QueryString("FNAME") & "%"
			cMI = "%" & Request.QueryString("MI") & "%"
			cSD = "%" & Request.QueryString("SD") & "%"
			cPhone = "%" & Request.QueryString("PHONE") & "%"
			cState = "%" & Request.QueryString("STATE") & "%"
			cCity = "%" & Request.QueryString("CITY") & "%"
			cZip = "%" & Request.QueryString("ZIP") & "%"
		Case "E"
			cLastName = Request.QueryString("LNAME")
			cFirstName = Request.QueryString("FNAME")
			cMI = Request.QueryString("MI")
			cSD = Request.QueryString("SD")
			cPhone = Request.QueryString("PHONE")
			cState = Request.QueryString("STATE")
			cCity = Request.QueryString("CITY")
			cZip = Request.QueryString("ZIP")
	End Select
	
	cPhone = Replace(cPhone, "'" , "''")
	cLastName = Replace(cLastName, "'" , "''")
	cFirstName = Replace(cFirstName, "'" , "''")
	cMI = Replace(cMI, "'" , "''")
	cCity = Replace(cCity, "'" , "''")
	cState = Replace(cState, "'" , "''")
	cZip = Replace(cZip, "'" , "''")
	
	cWhere = ""
	If Request.QueryString("LNAME") <> "" Then
		cWhere = "UPPER(NAME_LAST) LIKE '" & UCase(cLastName) & "'"
	End If
	If Request.QueryString("FNAME") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "UPPER(NAME_FIRST) LIKE '" & UCase(cFirstName) & "'"
	End If
	If Request.QueryString("MI") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "UPPER(NAME_MI) LIKE '" & UCase(cMI) & "'"
	End If
	If Request.QueryString("SD") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "SPECIFIC_DESTINATION_ID LIKE '" & cSD & "'"
	End If
	If Request.QueryString("PHONE") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "PHONE LIKE '" & cPhone & "'"
	End If
	If Request.QueryString("CITY") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "UPPER(CITY) LIKE '" & UCase(cCity) & "'"
	End If
	If Request.QueryString("STATE") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "UPPER(STATE) LIKE '" & UCase(cState) & "'"
	End If
	If Request.QueryString("ZIP") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "ZIP LIKE '" & cZip & "'"
	End If
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	oConn.Open ConnectionString

	if lFullSearch then
		cSQL = "SELECT * FROM SPECIFIC_DESTINATION" 
		If cWhere <> "" Then
			cSQL = cSQL & " WHERE (" & cWhere & ")" 
		end if
	else
		cSQL = "SELECT * FROM SPECIFIC_DESTINATION WHERE ACCNT_HRCY_STEP_ID=" & cCN
		If cWhere <> "" Then
			cSQL = cSQL & " AND (" & cWhere & ")" 
		end if
	End If
	cSQL = cSQL & " ORDER BY NAME_LAST" 
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.MaxRecords = MAXRECORDCOUNT
	oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
End If
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</HEAD>

<BODY  bgcolor='<%= BODYBGCOLOR %>'  leftmargin=2 topmargin=2 rightmargin=0 bottommargin=2>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%">
<thead CLASS="ResultHeader">
<TR>
<% if lFullSearch then%>
<td class="thd"><div id><nobr>AHS ID</div></td>
<%end if%>
<td class="thd"><div id><nobr>SD ID</div></td>
<td class="thd"><div id><nobr>Last Name</div></td>
<td class="thd"><div id><nobr>Mid. Init.</div></td>
<td class="thd"><div id><nobr>First Name</div></td>
<td class="thd"><div id><nobr>Title</div></TD>
<td class="thd"><div id><nobr>Address 1</div></TD>
<td class="thd"><div id><nobr>Address 2</div></TD>
<td class="thd"><div id><nobr>City</div></TD>
<td class="thd"><div id><nobr>State</div></TD>
<td class="thd"><div id><nobr>Zip Code</div></TD>
<td class="thd"><div id><nobr>Phone</div></TD>
</TR>
</THEAD>
<tbody ID="TableRows">
<%
nRecCount = 0
If Request.QueryString("SEARCHTYPE") <> "" Then
	Do While Not oRS.EOF 
		nRecCount = nRecCount + 1

%>
<TR ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" SD="<%=oRS.fields("SPECIFIC_DESTINATION_ID")%>">
<% if lFullSearch then%>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("ACCNT_HRCY_STEP_ID")) %></TD>
<%end if%>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("SPECIFIC_DESTINATION_ID")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("NAME_LAST")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("NAME_MI")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("NAME_FIRST")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("TITLE")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("ADDRESS1")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("ADDRESS2")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("CITY")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("STATE")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("ZIP")) %></TD>
<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("PHONE")) %></TD>
</TR>
<%
		oRS.movenext
	Loop
If oRS.EOF AND oRS.BOF Then 
%>
<TR>
<TD CLASS=LABEL COLSPAN=10 ALIGN="MIDDLE">No Specific Destinations Found</TD>
</TR>
<%
End If
oRS.Close
set oRS = nothing

oConn.Close
set oConn = nothing
End If 
%>
</tbody>
</TABLE>
</DIV>
</FIELDSET>
<SCRIPT LANGUAGE="VBScript">
<%	
If nRecCount > 0 Then %>
	if Parent.frames("TOP").document.readyState = "complete" then
		curCount = <%=nRecCount%>
		if curCount = <%=MAXRECORDCOUNT%> then
			Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
		else		
			Parent.frames("TOP").UpdateStatus("Record count is <%=nRecCount%>")
		end if		
	end if
<%	End If %>

Sub window_onload
<% If Request.Form <> "" Then %>
		if 0 < Document.all.tblResult.rows.length then
			call multiselect( Document.all.tblResult.rows(1))
		end if
<%End if %>
End Sub

Function GetSDID
	GetSDID = getmultipleindex(document.all.tblResult, "SD")
End Function

Sub document_onkeydown
	select case window.event.keycode
		case 8:
			window.event.keyCode = 0
			window.event.returnValue = 0
		case 38:
			call relativemultiselect( Document.all.tblResult, -1 )
		case 40:
			call relativemultiselect( Document.all.tblResult, 1 )
		case 13:
			i = getselectedindex( Document.all.tblResult )
			if 0 < i then
				dblhighlight(Document.all.tblResult.rows(i))
			end if
		case else:
	end select
End Sub
-->
</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "SD")
		return objRow.getAttribute("SD");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
		
}
</SCRIPT>

</BODY>
</HTML>
