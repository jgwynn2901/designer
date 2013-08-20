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
		Parent.frames("TOP").document.all.FID.value = ""
	end if
End Sub

Function GetFID
	GetFID = Trim(getmultipleindex(document.all.tblFields, "FID"))
End Function

Function GetFIDName
	GetFIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

Function GetFIDCaption
	GetFIDCaption = getmultipleindex(document.all.tblFields, "CAPTION")
End Function

Function GetFIDInputType
	GetFIDInputType = getmultipleindex(document.all.tblFields, "INPUTTYPE")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "FID")
		return objRow.getAttribute("FID");
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
			<td class="thd"><div id><nobr>Fee Type Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
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
				FID = Request.QueryString("SearchFID") & "%"
				NAME = Request.QueryString("SearchNAME") & "%"
			Case "C"
				FID = "%" & Request.QueryString("SearchFID") & "%"
				NAME = "%" & Request.QueryString("SearchNAME") & "%"
			Case "E"
				FID = Request.QueryString("SearchFID")
				NAME = Request.QueryString("SearchNAME")
		End Select
	
		FID = Replace(FID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		
		If Request.QueryString("SearchFID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(FEE_TYPE_ID) LIKE '" & UCASE(FID)  & "'"
		End If
		If Request.QueryString("SearchNAME") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT * FROM FEE_TYPE "
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS
			End If
			SQLST = SQLST & " ORDER BY NAME" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" FID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No fee types found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  FID='<%=RS("FEE_TYPE_ID")%>'>
	<td NOWRAP CLASS="ResultCell" ID="AID"><%=renderCell(RS("FEE_TYPE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
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

