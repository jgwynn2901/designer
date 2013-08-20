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
<title>Coverage Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.COVID.value = ""
	end if
End Sub

Function GetCOVID
	GetCOVID = getmultipleindex(document.all.tblFields, "COVID")
End Function
</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "COVID")
		return objRow.getAttribute("COVID");
}
</SCRIPT>
</head>
<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Coverage Id</div></td>
			<td class="thd"><div id><nobr>Policy Id</div></td>
			<td class="thd"><div id><nobr>Vehicle Id</div></td>
			<td class="thd"><div id><nobr>Effective Date</div></td>
			<td class="thd"><div id><nobr>Expiration Date</div></td>						
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
				COVID = Request.QueryString("SearchCOVID") & "%"
				POLICY_ID = Request.QueryString("SearchPOLICY_ID") & "%"
				VEHICLE_ID = Request.QueryString("SearchVEHICLE_ID") & "%"
				EFFECTIVE_DATE = Request.QueryString("SearchEFFECTIVE_DATE") & "%"
				EXPIRATION_DATE = Request.QueryString("SearchEXPIRATION_DATE") & "%"
			Case "C"
				COVID = "%" & Request.QueryString("SearchCOVID") & "%"
				POLICY_ID = "%" & Request.QueryString("SearchPOLICY_ID") & "%"
				VEHICLE_ID = "%" & Request.QueryString("SearchVEHICLE_ID") & "%"
				EFFECTIVE_DATE = "%" & Request.QueryString("SearchEFFECTIVE_DATE") & "%"
				EXPIRATION_DATE = "%" & Request.QueryString("SearchEXPIRATION_DATE") & "%"
			Case "E"
				COVID = Request.QueryString("SearchCOVID")
				POLICY_ID = Request.QueryString("SearchPOLICY_ID")
				VEHICLE_ID = Request.QueryString("SearchVEHICLE_ID")
				EFFECTIVE_DATE = Request.QueryString("SearchEFFECTIVE_DATE")
				EXPIRATION_DATE = Request.QueryString("SearchEXPIRATION_DATE")
		End Select
		COVID = Replace(COVID, "'", "''")
		POLICY_ID = Replace(POLICY_ID, "'", "''")
		VEHICLE_ID = Replace(VEHICLE_ID, "'", "''")
		EXPIRATION_DATE = Replace(EXPIRATION_DATE, "'", "''")
		EFFECTIVE_DATE = Replace(EFFECTIVE_DATE, "'", "''")
		If Request.QueryString("SearchPOLICY_ID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(POLICY_ID) LIKE '" & UCASE(POLICY_ID)  & "'"
		End If
		If Request.QueryString("SearchCOVID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "COVERAGE_ID LIKE '" & COVID & "'"
		End If
		If Request.QueryString("SearchVEHICLE_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(VEHICLE_ID) LIKE '" & UCASE(VEHICLE_ID) & "'"
		End If
		If Request.QueryString("SearchEFFECTIVE_DATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(EFFECTIVE_DATE) LIKE '" & UCASE(EFFECTIVE_DATE) & "'"
		End If
		If Request.QueryString("SearchEXPIRATION_DATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(EXPIRATION_DATE) LIKE '" & UCASE(EXPIRATION_DATE) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT COVERAGE_ID, POLICY_ID, VEHICLE_ID, EFFECTIVE_DATE, EXPIRATION_DATE FROM COVERAGE "
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS
			End If
			 SQLST = SQLST & " ORDER BY COVERAGE_ID" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" COVID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No coverage found re-check your criteria</td>
</tr>
<%		Else
			Do While Not RS.EOF
			RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  COVID='<%=RS("COVERAGE_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("COVERAGE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("POLICY_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("VEHICLE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("EFFECTIVE_DATE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("EXPIRATION_DATE"))%></td>
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
