<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vehicle Driver Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
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
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
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

	VID = CStr(Request.QueryString("VID"))
	If VID <> "NEW" And VID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM DRIVER WHERE VEHICLE_ID = " & VID 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" DID="<%=RS("DRIVER_ID")%>">
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DRIVER_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SSN"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME_FIRST"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME_LAST"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CITY"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ZIP"))%></td>
	</tr>

<%
		RS.MoveNext
		Loop
		
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing

	End If
%>

</tbody>
</table>
</div>
</BODY>
</HTML>


