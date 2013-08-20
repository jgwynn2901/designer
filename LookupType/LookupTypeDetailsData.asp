<%	Response.Expires = 0
	On Error Resume Next
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Lookup Type Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedLUCID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedLUCID = document.all.tblFields.rows(idx).getAttribute("LUCID")
	Else
		GetSelectedLUCID = ""
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
			<td class="thd"><div id><nobr>Lookup Id</div></td>
			<td class="thd"><div id><nobr>Code</div></td>
			<td class="thd"><div id><nobr>Value</div></td>
			<td class="thd"><div id><nobr>Sequence</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">

<%

	LUTID = CStr(Request.QueryString("LUTID"))
	If LUTID <> "NEW" And LUTID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM LU_CODE WHERE LU_TYPE_ID = " & LUTID & " ORDER BY SEQUENCE" 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" LUCID="<%=RS("LU_ID")%>">
	<td NOWRAP CLASS="ResultCell" ID="LUCID"><%=RS("LU_ID")%></td>
	<td NOWRAP CLASS="ResultCell" ID="CODE"><%=RS("CODE")%></td>
	<td NOWRAP CLASS="ResultCell" ID="VALUE"><%=RS("VALUE")%></td>
	<td NOWRAP CLASS="ResultCell" ID="SEQUENCE"><%=RS("SEQUENCE")%></td>
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


