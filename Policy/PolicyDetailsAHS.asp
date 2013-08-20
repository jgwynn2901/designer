<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<title>AHS Policy Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedCVID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedCVID = document.all.tblFields.rows(idx).getAttribute("CVID")
	Else
		GetSelectedCVID = ""
	End If
End Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" scrolling="auto">
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>AHS Policy Id</div></td>
			<td class="thd"><div id><nobr>AHSID</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Active St. Dt</div></td>
			<td class="thd"><div id><nobr>Active End Dt</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	PID = CStr(Request.QueryString("PID"))
	PROPID = CStr(Request.QueryString("PROPID"))
	VID = CStr(Request.QueryString("VID"))
If PROPID <> "NEW" AND VID <> "NEW" AND PID <> "NEW" Then
	If PID <> "" OR PROPID <> "" OR VID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM AHS_POLICY WHERE "
		If PROPID <> "" Then
			SQLPART = SQLPART & "PROPERTY_ID = " & PROPID 
		ElseIf VID <> "" Then
			SQLPART = SQLPART & "VEHICLE_ID = " & VID 
		ElseIf PID <> "" Then
			SQLPART = SQLPART & "POLICY_ID = " & PID 
		End If

		Set RS = Conn.Execute(SQLST & SQLPART)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" CVID="<%=RS("COVERAGE_ID")%>">
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("AHS_POLICY_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LOB_CD"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACTIVE_START_DT"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACTIVE_END_DT"))%></td>
	</tr>

<%
		RS.MoveNext
		Loop
		
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing

	End If
End If
%>

</tbody>
</table>
</div>
</BODY>
</HTML>


