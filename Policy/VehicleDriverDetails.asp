<%	Response.Expires = 0
	Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vehicle Driver details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedDID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedDID = document.all.tblFields.rows(idx).getAttribute("DID")
	Else
		GetSelectedDID = ""
	End If
End Function
</script>

</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" ID="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Driver ID</div></td>
			<td class="thd"><div id><nobr>SSN</div></td>
			<td class="thd"><div id><nobr>First Name</div></td>
			<td class="thd"><div id><nobr>Last Name</div></td>
			<td class="thd"><div id><nobr>City</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Zip</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
dim cPID, oConn, cSQL, oRS

cPID = CStr(Request.QueryString("PID"))
If cPID <> "NEW" And cPID <> "" Then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
		
	cSQL ="Select * From DRIVER Where POLICY_ID = " & cPID
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" DID="<%=oRS("DRIVER_ID")%>">
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DRIVER_ID"))%> </td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("SSN"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("NAME_FIRST"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("NAME_LAST"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("CITY"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("ZIP"))%></td>
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


