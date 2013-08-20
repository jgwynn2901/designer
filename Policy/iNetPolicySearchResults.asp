<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>iNetPolicy Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Function getRowID
	dim idx

	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		getRowID = document.all.tblFields.rows(idx).getAttribute("rowID")
	Else
		getRowID = ""
	End If
End Function

Function getKey
	getKey = getmultipleindex(document.all.tblFields, "rowID")
End Function

-->
</SCRIPT>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "rowID")
		return objRow.getAttribute("rowID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
		
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Client Code</div></td>
			<td class="thd"><div id><nobr>Policy ID</div></td>
			<td class="thd"><div id><nobr>Carrier Name</div></td>
			<td class="thd"><div id><nobr>Insured Name</div></td>
			<td class="thd"><div id><nobr>Address 1</div></td>
			<td class="thd"><div id><nobr>Address 2</div></td>
			<td class="thd"><div id><nobr>City</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Zip</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
	dim nRecCount, cClientCode, cPolicyID, cCarrierName, cInsuredName
	dim cSTATE, cZIP, oConn, cSQL, cWhereCls, oRS
	dim cKey
	
	nRecCount = -1
	If Request.QueryString("SEARCHTYPE") <> "" Then
		nRecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				cClientCode = Request.QueryString("ClientCode") & "%"
				cPolicyID = Request.QueryString("PolicyID") & "%"
				cCarrierName = Request.QueryString("CarrierName") & "%"
				cInsuredName = Request.QueryString("InsuredName") & "%"
				cSTATE = Request.QueryString("SearchState") & "%"
				cZIP = Request.QueryString("SearchZip") & "%"
			Case "C"
				cClientCode = "%" & Request.QueryString("ClientCode") & "%"
				cPolicyID = "%" & Request.QueryString("PolicyID") & "%"
				cCarrierName = "%" & Request.QueryString("CarrierName") & "%"
				cInsuredName = "%" & Request.QueryString("InsuredName") & "%"
				cSTATE = "%" & Request.QueryString("SearchState") & "%"
				cZIP = "%" & Request.QueryString("SearchZip") & "%"
			Case "E"
				cClientCode = Request.QueryString("ClientCode")
				cPolicyID = Request.QueryString("PolicyID")
				cCarrierName = Request.QueryString("CarrierName")
				cInsuredName = Request.QueryString("InsuredName")
				cSTATE = Request.QueryString("SearchState")
				cZIP = Request.QueryString("SearchZip")
		End Select
		cClientCode = Replace(cClientCode, "'", "''")
		cPolicyID = Replace(cPolicyID, "'", "''")
		cCarrierName = Replace(cCarrierName, "'", "''")
		cInsuredName = Replace(cInsuredName, "'", "''")
		cSTATE = Replace(cSTATE, "'", "''")
		cZIP = Replace(cZIP, "'", "''")
		If Request.QueryString("ClientCode") <> "" Then
			cWhereCls = "Client_CD LIKE '" & cClientCode & "'"
		End If
		If Request.QueryString("PolicyID") <> "" Then
			If cWhereCls <> "" Then 
				cWhereCls = cWhereCls & " AND "
			End If
			cWhereCls = cWhereCls & "POLICY_IDENTIFIER LIKE '" & cPolicyID & "'"
		End If
		If Request.QueryString("CarrierName") <> "" Then
			If cWhereCls <> "" Then 
				cWhereCls = cWhereCls & " AND "
			End If
			cWhereCls = cWhereCls & "UPPER(CARRIER_NAME) LIKE '" & UCase(cCarrierName) & "'"
		End If
		If Request.QueryString("InsuredName") <> "" Then
			If cWhereCls <> "" Then 
				cWhereCls = cWhereCls & " AND "
			End If
			cWhereCls = cWhereCls & "UPPER(INSURED_NAME) LIKE '" & UCase(cInsuredName) & "'"
		End If
		If Request.QueryString("SearchState") <> "" Then
			If cWhereCls <> "" Then 
				cWhereCls = cWhereCls & " AND "
			End If
			cWhereCls = cWhereCls & "ADDRESS_STATE LIKE '" & cSTATE & "'"
		End If
		If Request.QueryString("SearchZip") <> "" Then
			If cWhereCls <> "" Then 
				cWhereCls = cWhereCls & " AND "
			End If
			cWhereCls = cWhereCls & "ADDRESS_ZIP LIKE '" & cZIP & "'"
		End If		

		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		cSQL = "SELECT * " &_
					"FROM INETPOLICY"
		If cWhereCls <> "" Then
			cSQL = cSQL & " WHERE " & cWhereCls
		End If
		cSQL = cSQL & " ORDER BY CLIENT_CD" 
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.MaxRecords = MAXRECORDCOUNT
		oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
		if oRS.EOF And oRS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" AID=''>
	<td COLSPAN="9"  NOWRAP CLASS="ResultCell" >No records found: re-check your criteria</td>
</tr>

<%			Else
				Do While Not oRS.EOF
					nRecCount = nRecCount + 1
					cKey = oRS("CLIENT_CD") & "^" & oRS("POLICY_IDENTIFIER")
%>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" rowID='<%=cKey%>'>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("CLIENT_CD"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("POLICY_IDENTIFIER"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("CARRIER_NAME"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("INSURED_NAME"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ADDRESS_LINE1"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ADDRESS_LINE2"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ADDRESS_CITY"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ADDRESS_STATE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ADDRESS_ZIP"))%></td>
</tr>
<%
				oRS.MoveNext
				Loop
			End If
			oRS.Close
			Set oRS = Nothing
			oConn.Close
			Set oConn = Nothing
End If
%>

</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">
<%	If nRecCount >= 0 Then %>
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=nRecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=nRecCount%>")
	end if		
end if
<%	End If %>
</SCRIPT>
</body>
</html>
