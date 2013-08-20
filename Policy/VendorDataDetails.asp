<%	Response.Expires = 0
	Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vendor Designator Created by XREF</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'82%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all"  width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Vendor Design.</div></td>
			<td class="thd"><div id><nobr>Limit 1</div></td>
			<td class="thd"><div id><nobr>Limit 2</div></td>
			<td class="thd"><div id><nobr>Deductible 1</div></td>
			<td class="thd"><div id><nobr>Deductible 2</div></td>
			<td class="thd"><div id><nobr>Active Start Date</div></td>
			<td class="thd"><div id><nobr>Active End Date</div></td>
			<td class="thd"><div id><nobr>File_Transmission </div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
dim cVID, oConn, cSQL, oRS
	
cVID = Request.QueryString("VID")
if cVID <> "NEW" then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL =  "SELECT * FROM VENDOR_CODES WHERE VEHICLE_ID = " & cVID 
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
		<tr ID="FieldRow" CLASS="ResultRow">
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("VENDOR_DESIGNATOR"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LIMIT1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LIMIT2"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DEDUCTIBLE1"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DEDUCTIBLE2"))%></td>
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


