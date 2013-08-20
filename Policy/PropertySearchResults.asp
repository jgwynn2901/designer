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
<title>Property Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.PROPID.value = ""
	end if
End Sub

Function GetPROPID
	GetPROPID = getmultipleindex(document.all.tblFields, "PROPID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "PROPID")
		return objRow.getAttribute("PROPID");
		
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Property ID</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>Address</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Zip</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount
	RecCount = -1
	
	If Request.QueryString <> "" Then
	RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				PROPID = Request.QueryString("SearchPROPID") & "%"
				POLICY_ID = Request.QueryString("SearchPOLICY_ID") & "%"
				PROPERTY_DESCRIPTION = Request.QueryString("SearchPROPERTY_DESCRIPTION") & "%"
				ADDRESS = Request.QueryString("SearchADDRESS") & "%"
				CITY = Request.QueryString("SearchCITY") & "%"
				STATE = Request.QueryString("SearchSTATE") & "%"
				ZIP = Request.QueryString("SearchZIP") & "%"
			Case "C"
				PROPID = "%" & Request.QueryString("SearchPROPID") & "%"
				POLICY_ID =  "%" & Request.QueryString("SearchPOLICY_ID") & "%"
				PROPERTY_DESCRIPTION =  "%" & Request.QueryString("SearchPROPERTY_DESCRIPTION") & "%"
				ADDRESS =  "%" & Request.QueryString("SearchADDRESS") & "%"
				CITY =  "%" & Request.QueryString("SearchCITY") & "%"
				STATE =  "%" & Request.QueryString("SearchSTATE") & "%"
				ZIP =  "%" & Request.QueryString("SearchZIP") & "%"
			Case "E"
				PROPID = Request.QueryString("SearchPROPID")
				POLICY_ID = Request.QueryString("SearchPOLICY_ID")
				PROPERTY_DESCRIPTION = Request.QueryString("SearchPROPERTY_DESCRIPTION")
				ADDRESS = Request.QueryString("SearchADDRESS")
				CITY = Request.QueryString("SearchCITY")
				STATE = Request.QueryString("SearchSTATE")
				ZIP =  Request.QueryString("SearchZIP")
		End Select
		PROPID = Replace(PROPID,"'","''")
		POLICY_ID = Replace(POLICY_ID,"'","''")
		PROPERTY_DESCRIPTION = Replace(PROPERTY_DESCRIPTION,"'","''")
		ADDRESS = Replace(ADDRESS,"'","''")
		CITY = Replace(CITY,"'","''")
		STATE = Replace(STATE,"'","''")
		ZIP = Replace(ZIP,"'","''")
		If Request.QueryString("SearchPROPID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(PROPERTY_ID) LIKE '" & UCASE(PROPID)  & "'"
		End If
		If Request.QueryString("SearchPOLICY_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "POLICY_ID LIKE '" & POLICY_ID & "'"
		End If
		If Request.QueryString("SearchPROPERTY_DESCRIPTION") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(PROPERTY_DESCRIPTION) LIKE '" & UCASE(PROPERTY_DESCRIPTION) & "'"
		End If
		If Request.QueryString("SearchADDRESS") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ADDRESS1) LIKE '" & UCASE(ADDRESS) & "'"
		End If
		If Request.QueryString("SearchCITY") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CITY) LIKE '" & UCASE(CITY) & "'"
		End If
		If Request.QueryString("SearchSTATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(STATE) LIKE '" & UCASE(STATE) & "'"
		End If
		If Request.QueryString("SearchZIP") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ZIP) LIKE '" & UCASE(ZIP) & "'"
		End If
		
		if WHERECLS <> "" then
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = ""
			SQLST = "SELECT PROPERTY_ID, ADDRESS1, CITY, STATE, ZIP, PROPERTY_DESCRIPTION FROM PROPERTY WHERE " & WHERECLS & " ORDER BY POLICY_ID" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  PROPID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell" >No property found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  PROPID='<%=RS("PROPERTY_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PROPERTY_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PROPERTY_DESCRIPTION"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS1"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ZIP"))%></td>
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

