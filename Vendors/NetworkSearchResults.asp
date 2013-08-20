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
		Parent.frames("TOP").document.all.NID.value = ""
	end if
End Sub

Function GetNID
	GetNID = getmultipleindex(document.all.tblFields, "NID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "NID")
		return objRow.getAttribute("NID");
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'90%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0"  rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Network ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
	dim RecCount
	RecCount = -1
	WHERECLS = ""
	If Request.QueryString <> "" Then
		RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				NID = Request.QueryString("SearchNID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
			Case "C"
				NID = "%" & Request.QueryString("SearchNID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
			Case "E"
				NID = Request.QueryString("SearchNID")
				NAME = Request.QueryString("SearchName")
		End Select
		NAME = Replace(NAME, "'", "''")
	
		If Request.QueryString("SearchNID") <> "" Then
			WHERECLS = WHERECLS & "NETWORK_ID LIKE '" & NID  & "'"
		End If
		If Request.QueryString("SearchName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "NAME LIKE '" & NAME & "'"
		End If
		If WHERECLS <> "" Then 
			WHERECLS = WHERECLS & " AND "
		End If
	
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM NETWORK "
		if WHERECLS <> "" then SQLST = SQLST & " WHERE " & WHERECLS
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  BID='' >
	<td COLSPAN=8 NOWRAP CLASS="ResultCell" >No Networks found re-check your criteria</td>
</tr>
	
<%		Else
			Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  NID='<%=RS("NETWORK_ID")%>'>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NETWORK_ID"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME"))%></td>
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
