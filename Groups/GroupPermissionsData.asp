<%	Response.Expires = 0
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Group Permissions Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedACCID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedACCID = document.all.tblFields.rows(idx).getAttribute("ACCID")
	Else
		GetSelectedACCID = ""
	End If
End Function


</script>

<!--#include file="..\lib\tablecommon.inc"-->
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0">
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Function</div></td>
			<td class="thd"><div id><nobr>Access Type</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	GID = CStr(Request.QueryString("GID"))
	If GID <> "NEW" And GID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ACCESSPERMISSIONS, FUNCTION,ACCESSTYPE WHERE " &_
				"ACCESSPERMISSIONS.FUNCTION_ID = FUNCTION.FUNCTION_ID AND " &_ 
				"ACCESSPERMISSIONS.ACCESSTYPE_ID = ACCESSTYPE.ACCESSTYPE_ID AND " &_ 
				"GROUP_ID = " & GID  & " ORDER BY FUNCTION_NM, ACCESSTYPE"
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" ACCID="<%=RS("ACCESS_ID")%>">
	<td NOWRAP CLASS="ResultCell" ID="FUNCTION"><%=RS("FUNCTION_NM")%></td>
	<td NOWRAP CLASS="ResultCell" ID="ACCESSTYPE"><%=RS("ACCESSTYPE")%></td>
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
</body>
</html>


