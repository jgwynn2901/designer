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
<title>Output Page Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.OPID.value = ""
	end if
End Sub

Function GetOPID
	GetOPID = getmultipleindex(document.all.tblFields, "OPID")
End Function

Function GetOPIDName
	GetOPIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

Function GetOPIDCaption
	GetOPIDCaption = getmultipleindex(document.all.tblFields, "CAPTION")
End Function

Function GetOPIDInputType
	GetOPIDInputType = getmultipleindex(document.all.tblFields, "INPUTTYPE")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "OPID")
		return objRow.getAttribute("OPID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
	else if (whichCol == "CAPTION")		
		return objRow.cells("CAPTION").innerText;
	else if (whichCol == "INPUTTYPE")		
		return objRow.cells("INPUTTYPE").innerText;
		
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Outpt Page ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>BMP</div></td>
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
				OPID = Request.QueryString("SearchOPID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				ODID = Request.QueryString("SearchODID") & "%"
				OPBMP= Request.QueryString("SearchBMP")  & "%"
			Case "C"
				OPID = "%" & Request.QueryString("SearchOPID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
				ODID = "%" & Request.QueryString("SearchODID") & "%"
				OPBMP= "%" & Request.QueryString("SearchBMP")  & "%"				
			Case "E"
				OPID = Request.QueryString("SearchOPID")
				NAME = Request.QueryString("SearchName")
				ODID = Request.QueryString("SearchODID")
				OPBMP= Request.QueryString("SearchBMP")
		End Select
		OPID = Replace(OPID,"'","''")
		NAME = Replace(NAME,"'","''")
		ODID = Replace(ODID,"'","''")
		OPBMP= Replace(OPBMP,"'","''")
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchOPID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "OUTPUT_PAGE_ID LIKE '" & OPID & "'"
		End If
		If Request.QueryString("SearchODID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(OUTPUTDEF_ID) LIKE '" & UCASE(ODID) & "'"
		End If
		If Request.QueryString("SearchBMP") <> "" Then
			If WHERECLS <> "" Then
				WHERECLS = WHERECLS & " AND "
			End if
			WHERECLS = WHERECLS & "UPPER(BACKGROUND_BMP) LIKE '" & UCASE(OPBMP) & "'"
		End If
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = SQLST & "SELECT NAME,OUTPUT_PAGE_ID, OUTPUTDEF_ID, BACKGROUND_BMP FROM OUTPUT_PAGE "
		If WHERECLS <> "" Then
			SQLST = SQLST & "WHERE " & WHERECLS 
		End If
		SQLST = SQLST & " ORDER BY NAME" 
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OPID='' >
	<td COLSPAN=4 NOWRAP CLASS="ResultCell">No Output Pages found re-check your criteria</td>
</tr>
	
	<%	Else
			Do While Not RS.EOF
				RecCount = RecCount + 1 %>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  OPID='<%=RS("OUTPUT_PAGE_ID")%>'>
	<td NOWRAP CLASS="ResultCell" ID="OPID"><%=renderCell(RS("OUTPUT_PAGE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("BACKGROUND_BMP"))%></td>
	</tr>

<%				RS.MoveNext
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
