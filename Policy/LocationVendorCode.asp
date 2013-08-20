<%	Response.Expires = 0
	Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vendor Designators Created by XREF</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all"  width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Loc. Cov. ID</div></td>
			<td class="thd"><div id><nobr>Vendor Design.</div></td>
			<td class="thd"><div id><nobr>Limit 1</div></td>
			<td class="thd"><div id><nobr>Deductible 1</div></td>
			<td class="thd"><div id><nobr>Active Start Date</div></td>
			<td class="thd"><div id><nobr>Active End Date</div></td>
			<td class="thd"><div id><nobr>File transmission Log ID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
dim oConn, oRS, cSQL, cAHS_ID

cAHS_ID = Request.QueryString("AHS_ID")
if cAHS_ID <> "NEW" then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "SELECT VC.* " & _
		"FROM ACCOUNT_HIERARCHY_STEP AHS, AHS_POLICY AHSP, VENDOR_CODES VC " & _
		"WHERE AHS.ACCNT_HRCY_STEP_ID = AHSP.ACCNT_HRCY_STEP_ID AND VC.AHS_POLICY_ID = AHSP.AHS_POLICY_ID " & _
		"AND AHS.ACCNT_HRCY_STEP_ID = " & cAHS_ID
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
		<tr ID="FieldRow" CLASS="ResultRow">
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("VENDOR_DESIGNATOR"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LIMIT1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DEDUCTIBLE1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ACTIVE_START_DATE"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ACTIVE_END_DATE"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("FILE_TRANSMISSION_LOG_ID"))%></td>
		</tr>
	<%
		oRS.MoveNext
	Loop
	oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
end if	
%>
</tbody>
</table>
</div>
</BODY>
</HTML>


