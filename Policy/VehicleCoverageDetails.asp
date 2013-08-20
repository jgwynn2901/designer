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

<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'82%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all"  width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Veh. Cov. ID</div></td>
			<td class="thd"><div id><nobr>Coverage Type</div></td>			
			<td class="thd"><div id><nobr>Limit 1</div></td>
			<td class="thd"><div id><nobr>Limit 2</div></td>
            <td class="thd"><div id><nobr>Deductible 1</div></td>
            <td class="thd"><div id><nobr>Deductible 2</div></td>
            <td class="thd"><div id><nobr>Active Start date</div></td>
            <td class="thd"><div id><nobr>Active End date</div></td>
            <td class="thd"><div id><nobr>File Trans. LogID</div></td>
			<td class="thd"><div id><nobr>Upload Key</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
dim cVID, oConn, oRS, cSQL

cVID = Request.QueryString("VID")
if cVID <> "NEW" then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL= "select * from VEHICLE_COVERAGE WHERE VEHICLE_ID =" & cVID _
		& " ORDER BY VEHICLE_COVERAGE_ID" 		
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
		<tr ID="FieldRow" CLASS="ResultRow">
		<td NOWRAP CLASS="ResultCell"><%=oRS("VEHICLE_COVERAGE_ID")%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("COVERAGE_TYPE"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LIMIT1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LIMIT2"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DEDUCTIBLE1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DEDUCTIBLE2"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ACTIVE_START_DT"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ACTIVE_END_DT"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=oRS("FILE_TRANSMISSION_LOG_ID")%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("UPLOAD_KEY"))%></td>
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


