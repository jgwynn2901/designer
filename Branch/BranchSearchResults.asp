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
<title>Branch Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.BID.value = ""
	end if
End Sub

Function GetBID
	GetBID = getmultipleindex(document.all.tblFields, "BID")
End Function

Function GetBIDOfficeName
	GetBIDOfficeName = getmultipleindex(document.all.tblFields, "OFFICENAME")
End Function

Function GetBNUM
	GetBNUM = getmultipleindex(document.all.tblFields, "BRANCHNUMBER")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "BID")
		return objRow.getAttribute("BID");
	else if (whichCol == "OFFICENAME")		
		return objRow.cells("OFFICENAME").innerText;
	else if (whichCol == "BRANCHNUMBER")		
		return objRow.cells("BRANCHNUMBER").innerText;
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0"  rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Branch ID</div></td>
			<td class="thd"><div id><nobr>Branch Number</div></td>
			<td class="thd"><div id><nobr>AH Load ID</div></td>
			<!--td class="thd"><div id><nobr>Office Number</div></td>
			<td class="thd"><div id><nobr>Status</div></td>			
			<td class="thd"><div id><nobr>Office Type</div></td-->			
			<td class="thd"><div id><nobr>Office Name</div></td>			
			<td class="thd"><div id><nobr>Address</div></td>			
			<td class="thd"><div id><nobr>City</div></td>			
			<td class="thd"><div id><nobr>State</div></td>			
			<td class="thd"><div id><nobr>Zip</div></td>			
			<td class="thd"><div id><nobr>Branch Type</div></td>			
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
				BID = Request.QueryString("SearchBID") & "%"
				BRANCHNUMBER = Request.QueryString("SearchBranchNumber") & "%"
				AHLOADID = Request.QueryString("SearchAHLoadID") & "%"
				STATUS = Request.QueryString("SearchStatus") & "%"
				OFFICENUMBER = Request.QueryString("SearchOfficeNumber") & "%"
				OFFICETYPE = Request.QueryString("SearchOfficeType") & "%"
				OFFICENAME = Request.QueryString("SearchOfficeName") & "%"
				ADDRESS = Request.QueryString("SearchAddress") & "%"
				CITY = Request.QueryString("SearchCity") & "%"
				STATE = Request.QueryString("SearchState") & "%"
				ZIP = Request.QueryString("SearchZip") & "%"
				BRANCHTYPE = Request.QueryString("SearchBranchType") & "%"
			Case "C"
				BID = "%" & Request.QueryString("SearchBID") & "%"
				BRANCHNUMBER = "%" & Request.QueryString("SearchBranchNumber") & "%"
				AHLOADID = "%" & Request.QueryString("SearchAHLoadID") & "%"
				STATUS = "%" & Request.QueryString("SearchStatus") & "%"
				OFFICENUMBER = "%" & Request.QueryString("SearchOfficeNumber") & "%"
				OFFICETYPE = "%" & Request.QueryString("SearchOfficeType") & "%"
				OFFICENAME = "%" & Request.QueryString("SearchOfficeName") & "%"
				ADDRESS = "%" & Request.QueryString("SearchAddress") & "%"
				CITY = "%" & Request.QueryString("SearchCity") & "%"
				STATE = "%" & Request.QueryString("SearchState") & "%"
				ZIP = "%" & Request.QueryString("SearchZip") & "%"
				BRANCHTYPE = Request.QueryString("SearchBranchType") & "%"
			Case "E"
				BID = Request.QueryString("SearchBID")
				BRANCHNUMBER = Request.QueryString("SearchBranchNumber")
				AHLOADID = Request.QueryString("SearchAHLoadID")
				STATUS = Request.QueryString("SearchStatus")
				OFFICENUMBER = Request.QueryString("SearchOfficeNumber")
				OFFICETYPE = Request.QueryString("SearchOfficeType")
				OFFICENAME = Request.QueryString("SearchOfficeName")
				ADDRESS = Request.QueryString("SearchAddress")
				CITY = Request.QueryString("SearchCity")
				STATE = Request.QueryString("SearchState")
				ZIP = Request.QueryString("SearchZip")
				BRANCHTYPE = Request.QueryString("SearchBranchType") & "%"
		End Select
		BID = Replace(BID, "'", "''")
		BRANCHNUMBER = Replace(BRANCHNUMBER, "'", "''")
		AHLOADID = Replace(AHLOADID, "'", "''")
		STATUS = Replace(STATUS, "'", "''")
		OFFICENUMBER = Replace(OFFICENUMBER, "'", "''")
		OFFICETYPE = Replace(OFFICETYPE, "'", "''")
		OFFICENAME = Replace(OFFICENAME, "'", "''")
		OFFICENAME = Replace(OFFICENAME, "'", "''")
		ADDRESS = Replace(ADDRESS, "'", "''")
		CITY = Replace(CITY, "'", "''")
		STATE = Replace(STATE, "'", "''")
		ZIP = Replace(ZIP, "'", "''")
		BRANCHTYPE = Replace(BRANCHTYPE, "'", "''")
	
		If Request.QueryString("SearchBID") <> "" Then
			WHERECLS = WHERECLS & "BRANCH_ID LIKE '" & BID  & "'"
		End If
		If Request.QueryString("SearchBranchNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "BRANCH_NUMBER LIKE '" & BRANCHNUMBER & "'"
		End If
		If Request.QueryString("SearchAHLoadID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCOUNT_HIERARCHY_LOAD_ID LIKE '" & AHLOADID & "'"
		End If
		If Request.QueryString("SearchStatus") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(STATUS) LIKE '" & UCASE(STATUS) & "'"
		End If
		If Request.QueryString("SearchOfficeNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "OFFICE_NUMBER LIKE '" & OFFICENUMBER & "'"
		End If
		If Request.QueryString("SearchOfficeType") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(OFFICE_TYPE) LIKE '" & UCASE(OFFICETYPE) & "'"
		End If
		If Request.QueryString("SearchOfficeName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(OFFICE_NAME) LIKE '" & UCASE(OFFICENAME) & "'"
		End If		
		If Request.QueryString("SearchAddress") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(ADDRESS_1) LIKE '" & UCASE(ADDRESS) & "'"
		End If		
		If Request.QueryString("SearchCity") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CITY) LIKE '" & UCASE(CITY) & "'"
		End If
		If Request.QueryString("SearchState") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "STATE LIKE '" & STATE & "'"
		End If
		If Request.QueryString("SearchZip") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ZIP LIKE '" & ZIP & "'"
		End If		
	

		If Request.QueryString("BranchTypeFilter") <> "" Then 
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "BRANCH_TYPE LIKE '" & Request.QueryString("BranchTypeFilter") & "'"
		Else
			If Request.QueryString("SearchBranchType") <> "" Then
				If WHERECLS <> "" Then 
					WHERECLS = WHERECLS & " AND "
				End If
				WHERECLS = WHERECLS & "BRANCH_TYPE LIKE '" & BRANCHTYPE & "'"
			End If
		End If

			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM BRANCH "
		if WHERECLS <> "" then SQLST = SQLST & " WHERE " & WHERECLS
		SQLST = SQLST & " ORDER BY BRANCH_ID" 
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  BID='' >
	<td COLSPAN=8 NOWRAP CLASS="ResultCell" >No branches found re-check your criteria</td>
</tr>
	
<%		Else
			Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  BID='<%=RS("BRANCH_ID")%>'>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("BRANCH_ID"))%></td>
		<td NOWRAP CLASS="ResultCell" ID="BRANCHNUMBER"><%=renderCell(RS("BRANCH_NUMBER"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCOUNT_HIERARCHY_LOAD_ID"))%></td>
		<!--td NOWRAP CLASS="ResultCell"><%=renderCell(RS("OFFICE_NUMBER"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATUS"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("OFFICE_TYPE"))%></td-->
		<td NOWRAP CLASS="ResultCell" ID="OFFICENAME"><%=renderCell(RS("OFFICE_NAME"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS_1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CITY"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATE"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ZIP"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("BRANCH_TYPE"))%></td>
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
