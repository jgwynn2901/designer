<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
	
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Accounts Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Function GetSelectedAHSID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedAHSID = document.all.tblFields.rows(idx).getAttribute("AHSID")
	Else
		GetSelectedAHSID = ""
	End If
End Function

Function GetAHSID

	GetAHSID = getmultipleindex(document.all.tblFields, "AHSID")
End Function

</script>

<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Account ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	UID = CStr(Request.QueryString("UID"))
	If UID <> "NEW" And UID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT /*+ INDEX_COMBINE(AU) */ A.ACCNT_HRCY_STEP_ID, " &_
		 "A.NAME "&_
		"FROM ACCOUNT_HIERARCHY_STEP A,  "&_
		"ACCOUNT_USER AU  "&_
		"WHERE A.ACCNT_HRCY_STEP_ID = AU.ACCNT_HRCY_STEP_ID + 0  "&_
		"AND AU.USER_ID = " & UID &_
		" ORDER BY AU.ACCNT_HRCY_STEP_ID "
		
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" AHSID="<%=RS("ACCNT_HRCY_STEP_ID")%>">
	<td NOWRAP CLASS="ResultCell"><%=RS("ACCNT_HRCY_STEP_ID")%></td>
	<td NOWRAP CLASS="ResultCell"><%=RS("NAME")%></td>
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


