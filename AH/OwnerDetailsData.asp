<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next

%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>AHS Owner </title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">

Function GetSelectedAHSOID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedAHSOID = document.all.tblFields.rows(idx).getAttribute("AHSOID")
	Else
		GetSelectedAHSOID = ""
	End If
End Function

Function GetSelectedAHSID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedAHSID = document.all.tblFields.rows(idx).getAttribute("AHSID")
	Else
		GetSelectedAHSID = ""
	End If
End Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>

<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" rules=all  cellSpacing="0" scrolling="auto" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Owner ID</div></td>			
			<td class="thd"><div id><nobr>AHS Owner ID</div></td>
			<td class="thd"><div id><nobr>AHS ID</div></td>			
			<td class="thd"><div id><nobr>Active Start Date</div></td>			
			<td class="thd"><div id><nobr>Active End Date</div></td>			
			<td class="thd"><div id><nobr>File Transmission Log ID</div></td>
			
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	OID = CStr(Request.QueryString("searchOID"))
	If OID <> "NEW" And OID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * from AHS_OWNER  " &_
				" WHERE OWNER_ID  = " & OID & "  ORDER BY OWNER_ID" 
				
		Set RS = Conn.Execute(SQLST)
		'REsponse.write(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>

	<tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);"  AHSOID="<%=RS("AHS_OWNER_ID")%>" AHSID="<%=RS("ACCNT_HRCY_STEP_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("OWNER_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("AHS_OWNER_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("accnt_hrcy_step_id"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ACTIVE_START_DT"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ACTIVE_END_DT"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("FILE_TRANSMISSION_LOG_ID"))%></td>
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


