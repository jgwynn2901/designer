<%	
Response.Expires = 0
Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Location Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function getSelectedAHS_ID
	dim nIdx
	
	nIdx = CInt(getselectedindex(document.all.tblFields))
	If nIdx <> -1 Then
		getSelectedAHS_ID = document.all.tblFields.rows(nIdx).getAttribute("AHS_ID")
	Else
		getSelectedAHS_ID = ""
	End If
End Function

Function getSelectedAHS_POL_ID
	dim nIdx
	
	nIdx = CInt(getselectedindex(document.all.tblFields))
	If nIdx <> -1 Then
		getSelectedAHS_POL_ID = document.all.tblFields.rows(nIdx).getAttribute("AHS_POL_ID")
	Else
		getSelectedAHS_POL_ID = ""
	End If
End Function

</script>

<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>AHS ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Address</div></td>
			<td class="thd"><div id><nobr>City</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>ZIP</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
dim cPID, oConn, oRS, cSQL, cLOB

cPID = Request.QueryString("PID")
cLOB = Request.QueryString("LOB")
If cPID <> "NEW" And cPID <> "" Then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "SELECT AHS.*, AHSP.AHS_POLICY_ID " _
		& "FROM ACCOUNT_HIERARCHY_STEP AHS, AHS_POLICY AHSP " _
		& "WHERE AHS.ACCNT_HRCY_STEP_ID = AHSP.ACCNT_HRCY_STEP_ID " _
		& "AND AHSP.POLICY_ID = " & cPID _
		& " AND AHSP.LOB_CD = '" & cLOB & "'"

	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" AHS_POL_ID="<%=oRS("AHS_POLICY_ID")%>" AHS_ID="<%=oRS("ACCNT_HRCY_STEP_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("ADDRESS_1"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("CITY"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("ZIP"))%></td>
	</tr>

<%
		oRS.MoveNext
	Loop
	oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
End If
%>
</tbody>
</table>
</div>
</BODY>
</HTML>


