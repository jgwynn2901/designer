<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Jurisdiction State Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<script language="VBScript">
Function GetSelectedState
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedState = document.all.tblFields.rows(idx).getAttribute("STATE")
	Else
		GetSelectedState = ""
	End If
End Function
</script>

</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0"  scrolling="auto">
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" scrolling="auto">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>State</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	PID = CStr(Request.QueryString("PID"))
	If PID <> "NEW" And PID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM JURISDICTION_STATE WHERE POLICY_ID = " & PID 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" STATE="<%=RS("STATE")%>">
	<td NOWRAP CLASS="ResultCell"><%=RS("STATE")%></td>
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


