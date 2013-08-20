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
<title>Attribute Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.COID.value = ""
	end if
End Sub

Function GetCOID
	GetCOID = getmultipleindex(document.all.tblFields, "COID")
End Function

Function GetCOIDName
	GetCOIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function


</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "COID")
		return objRow.getAttribute("COID");
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
			<td class="thd"><div id><nobr>Contact Id</div></td>
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
				COID = Request.QueryString("SearchCOID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				sTYPE = Request.QueryString("SearchsTYPE") & "%"
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION") & "%"
			Case "C"
				COID = "%" & Request.QueryString("SearchCOID") & "%"
				NAME = "%" & Request.QueryString("SearchNAME") & "%"
				sTYPE = "%" & Request.QueryString("SearchsTYPE") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDESCRIPTION") & "%"
			Case "E"
				COID = Request.QueryString("SearchAID")
				NAME = Request.QueryString("SearchName")
				sTYPE = Request.QueryString("SearchsTYPE")
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION")
		End Select
		COID = Replace(COID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		sTYPE = Replace(sTYPE, "'", "''")
		DESCRIPTION = Replace(DESCRIPTION, "'", "''")
	
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchCOID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "CONTACT_ID LIKE '" & COID & "'"
		End If
		If Request.QueryString("SearchsTYPE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(TYPE) LIKE '" & UCASE(sTYPE) & "'"
		End If
		If Request.QueryString("SearchDESCRIPTION") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT NAME, DESCRIPTION, CONTACT_ID FROM CONTACT "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS 
			End If
			
			SQLST = SQLST & " ORDER BY NAME" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  COID='' >
	<td COLSPAN=10 NOWRAP CLASS="ResultCell">No Contacts found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  COID='<%=RS("CONTACT_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CONTACT_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
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
