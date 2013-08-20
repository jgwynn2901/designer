<%	Response.Expires = 0
	Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vendor Designator Created by XREF</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0"  width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Vendor Designator</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	VID = CStr(Request.QueryString("VID"))
	If VID <> "NEW" And VID <> "" Then
	   Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		 SQLST = "SELECT VC.VENDOR_DESIGNATOR"
         SQLST=SQLST & "FROM POLICY P,VEHICLE V,VENDOR_CODES VC"
         SQLST=SQLST & "WHERE VC.VEHICLE_ID(+) = V.VEHICLE_ID"
         SQLST=SQLST & "AND P.POLICY_ID = V.POLICY_ID(+)"
         SQLST=SQLST & " AND VC.VEHICLE_ID= " & VID 

		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	
	<td NOWRAP CLASS="ResultCell"><%=RS("VENDOR_DESIGNATOR")%></td>
	
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


