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
<title>Attribute Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.AID.value = ""
	end if
End Sub

Function GetAID
	GetAID = getmultipleindex(document.all.tblFields, "AID")
End Function

Function GetAIDName
	GetAIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

Function GetAIDCaption
	GetAIDCaption = getmultipleindex(document.all.tblFields, "CAPTION")
End Function

Function GetAIDInputType
	GetAIDInputType = getmultipleindex(document.all.tblFields, "INPUTTYPE")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "AID")
		return objRow.getAttribute("AID");
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
			<td class="thd"><div id><nobr>Attribute Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Caption</div></td>
			<td class="thd" style="display:none"><div id><nobr>Input Type</div></td>			
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
				AID = Request.QueryString("SearchAID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				CAPTION = Request.QueryString("SearchCaption") & "%"
				DESCRIPTION = Request.QueryString("SearchDescription") & "%"
				HELP = Request.QueryString("SearchHelpString") & "%"
				INPUTTYPE = Request.QueryString("SearchInputType") & "%"
			Case "C"
				AID = "%" & Request.QueryString("SearchAID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
				CAPTION = "%" & Request.QueryString("SearchCaption") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDescription") & "%"
				HELP = "%" & Request.QueryString("SearchHelpString") & "%"
				INPUTTYPE = "%" & Request.QueryString("SearchInputType") & "%"
			Case "E"
				AID = Request.QueryString("SearchAID")
				NAME = Request.QueryString("SearchName")
				CAPTION = Request.QueryString("SearchCaption")
				DESCRIPTION = Request.QueryString("SearchDescription")
				HELP = Request.QueryString("SearchHelpString")
				INPUTTYPE = Request.QueryString("SearchInputType")
		End Select
	
		AID = Replace(AID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		CAPTION = Replace(CAPTION, "'", "''")
		DESCRIPTION = Replace(DESCRIPTION, "'", "''")
		HELP = Replace(HELP, "'", "''")
		INPUTTYPE = Replace(INPUTTYPE, "'", "''")
		
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchAID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ATTRIBUTE_ID LIKE '" & AID & "'"
		End If
		If Request.QueryString("SearchCaption") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CAPTION) LIKE '" & UCASE(CAPTION) & "'"
		End If
		If Request.QueryString("SearchDescription") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		If Request.QueryString("SearchHelpString") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(HELPSTRING) LIKE '" & UCASE(HELP) & "'"
		End If
		If Request.QueryString("SearchInputType") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(INPUTTYPE) LIKE '" & UCASE(INPUTTYPE) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT NAME,CAPTION, INPUTTYPE, ATTRIBUTE_ID FROM ATTRIBUTE "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS 
			End If
			SQLST = SQLST & " ORDER BY NAME" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
			if RS.EOF And RS.BOF then %>
<tr ID="FieldRow" CLASS="ResultRow" AID='' >
	<td COLSPAN=4 NOWRAP CLASS="ResultCell">No attributes found re-check your criteria</td>
</tr>
<%		Else
			Do While Not RS.EOF
			RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  AID='<%=RS("ATTRIBUTE_ID")%>'>
	<td NOWRAP CLASS="ResultCell" ID="AID"><%=renderCell(RS("ATTRIBUTE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("CAPTION"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="INPUTTYPE" style="display:none"><%=renderCell(RS("INPUTTYPE"))%></td>
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
