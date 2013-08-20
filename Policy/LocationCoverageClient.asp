<%	Response.Expires = 0
	Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Coverage Code Received From Client</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="left" style="WIDTH: 100%; HEIGHT: 100%">
<table cellPadding="2" cellSpacing="0" rules="all"  width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div><nobr>Loc. Cov. ID</div></NOBR></td>
			<td class="thd"><div><nobr>Coverage Type</div></NOBR></td>
			<td class="thd"><div><nobr>Limit 1</div></NOBR></td>
			<td class="thd"><div><nobr>Deductible 1</div></NOBR></td>
			<td class="thd"><div><nobr>Active Start Date</div></NOBR></td>
			<td class="thd"><div><nobr>Active End Date</div></NOBR></td>
			<td class="thd"><div><nobr>File transmission Log ID</div></NOBR></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
dim oConn, cSQL, oRS, cAHS_ID

cAHS_ID = Request.QueryString("AHS_ID")
if len(cAHS_ID) <> 0 then
	if cAHS_ID <> "NEW" then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		
		cSQL = "SELECT LC.* " & _
			"FROM ACCOUNT_HIERARCHY_STEP AHS, AHS_POLICY AHSP, LOCATION_COVERAGE LC " & _
			"WHERE AHS.ACCNT_HRCY_STEP_ID = AHSP.ACCNT_HRCY_STEP_ID AND LC.AHS_POLICY_ID = AHSP.AHS_POLICY_ID " & _
			"AND AHS.ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
			" ORDER BY LOCATION_COVERAGE_ID"
		Set oRS = oConn.Execute(cSQL)
		Do While Not oRS.EOF 
		%>
			<tr ID="FieldRow" CLASS="ResultRow">
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LOCATION_COVERAGE_ID"))%></td>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("COVERAGE_TYPE"))%></td>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LIMIT1"))%></td>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DEDUCTIBLE1"))%></td>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ACTIVE_START_DT"))%></td>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ACTIVE_END_DT"))%></td>
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
end if	
%>
</tbody>
</table>
</div>
</BODY>
</HTML>


