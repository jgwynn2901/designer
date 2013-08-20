<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Property Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedPROPID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedPROPID = document.all.tblFields.rows(idx).getAttribute("PROPID")
	Else
		GetSelectedPROPID = ""
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
			<td class="thd"><div id><nobr>Property ID</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>Location Description</div></td>
			<td class="thd"><div id><nobr>Address 1</div></td>
			<td class="thd"><div id><nobr>Address 2</div></td>
			<td class="thd"><div id><nobr>City</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>Zip</div></td>
			<td class="thd"><div id><nobr>Phone</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	PID = CStr(Request.QueryString("PID"))
	If PID <> "NEW" And PID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM PROPERTY WHERE POLICY_ID = " & PID 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" PROPID="<%=RS("PROPERTY_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("PROPERTY_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("PROPERTY_DESCRIPTION"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("PROPERTY_LOCATION_DESCRIPTION"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ADDRESS1"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ADDRESS2"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("CITY"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("STATE"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ZIP"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("PHONE"))%></td>
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


