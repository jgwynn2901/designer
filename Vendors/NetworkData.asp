<%	Response.Expires = 0
	Response.Buffer = true
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Network Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedVID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedVID = document.all.tblFields.rows(idx).getAttribute("VID")
	Else
		GetSelectedVID = ""
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
			<td class="thd"><div id><nobr>Vendor ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Enabled</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	NID = CStr(Request.QueryString("NID"))
	If NID <> "NEW" And NID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT V.VENDOR_ID, V.NAME, V.ENABLED_FLG FROM VENDOR_NETWORK VN ,VENDOR V WHERE VN.NETWORK_ID = " & NID & " AND V.VENDOR_ID = VN.VENDOR_ID"
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" VID="<%=RS("VENDOR_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("VENDOR_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ENABLED_FLG"))%></td>
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


