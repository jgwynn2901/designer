<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Permissions Data</title>
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

Function IsSelectedUserLevel
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		 curText = document.all.tblFields.rows(idx).cells("ACCESSTHROUGH").innerText
		 If Right(curText,3) = "(G)" Then
			IsSelectedUserLevel = false
		Else
			IsSelectedUserLevel = true
		End If
	Else
		IsSelectedUserLevel = false
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
			<td class="thd"><div id><nobr>Function</div></td>
			<td class="thd"><div id><nobr>Access Type</div></td>
			<td class="thd"><div id><nobr>Access Through</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	UID = CStr(Request.QueryString("UID"))
	If UID <> "NEW" And UID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ACCESSPERMISSIONS_VIEW WHERE USER_ID = " & UID  & " ORDER BY FUNCTION_NM, ACCESSTYPE, ACCESS_THROUGH"
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" ACCID="<%=RS("ACCESS_ID")%>">
	<td NOWRAP CLASS="ResultCell" ID="FUNCTION"><%=renderCell(RS("FUNCTION_NM"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="ACCESSTYPE"><%=renderCell(RS("ACCESSTYPE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="ACCESSTHROUGH"><%=renderCell(RS("ACCESS_THROUGH"))%></td>
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


