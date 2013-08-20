<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Group Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedGID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedGID = document.all.tblFields.rows(idx).getAttribute("GID")
	Else
		GetSelectedGID = ""
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
			<td class="thd"><div id><nobr>Group Id</div></td>
			<td class="thd"><div id><nobr>Group Name</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	UID = CStr(Request.QueryString("UID"))
	If UID <> "NEW" And UID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM GROUPS, USER_GROUP WHERE GROUPS.GROUP_ID = USER_GROUP.GROUP_ID(+) AND USER_GROUP.USER_ID = " & UID 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" GID="<%=RS("GROUP_ID")%>">
	<td NOWRAP CLASS="ResultCell"><%=RS("GROUP_ID")%></td>
	<td NOWRAP CLASS="ResultCell"><%=RS("GROUP_NM")%></td>
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


