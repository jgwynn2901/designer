<%	Response.Expires = 0
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Group Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedUID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedUID = document.all.tblFields.rows(idx).getAttribute("UID")
	Else
		GetSelectedUID = ""
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
			<td class="thd"><div id><nobr>User Id</div></td>
			<td class="thd"><div id><nobr>User Name</div></td>
			<td class="thd"><div id><nobr>Site</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	GID = CStr(Request.QueryString("GID"))
	If GID <> "NEW" And GID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM USERS_SITE_VIEW, USER_GROUP, GROUPS WHERE USERS_SITE_VIEW.USER_ID = USER_GROUP.USER_ID AND USER_GROUP.GROUP_ID = GROUPS.GROUP_ID AND GROUPS.GROUP_ID = " & GID 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" UID="<%=RS("USER_ID")%>">
	<td NOWRAP CLASS="ResultCell"><%=RS("USER_ID")%></td>
	<td NOWRAP CLASS="ResultCell"><%=RS("NAME")%></td>
	<td NOWRAP CLASS="ResultCell"><%=RS("SITE_NAME")%></td>
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


