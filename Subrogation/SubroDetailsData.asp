<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next

	BranchTextLen = 25
	RuleTextLen = 25
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Subrogation Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
function f_LastBARuleRecord
	if document.all.tblFields.Rows.Length <= 2 Then
		f_LastBARuleRecord = true
	else
		f_LastBARuleRecord = false
	end if
end Function

Function GetSelectedSRID
	dim idx
	
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedSRID = document.all.tblFields.rows(idx).getAttribute("SRID")
	Else
		GetSelectedSRID = ""
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
		    <td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Score</div></td>
			<td class="thd"><div id><nobr>Rule ID</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	STID = CStr(Request.QueryString("STID"))
	If STID <> "NEW" And STID <> "" Then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		SQLST = "SELECT * FROM SUBROGATION_DETECTION_RULE "
		SQLST = SQLST & "WHERE SUBROGATION_DETECTION_TYPE_ID = "& STID 
		SQLST = SQLST & " ORDER BY NAME" 
		Set oRS = oConn.Execute(SQLST)
		Do While Not oRS.EOF
%>
    <tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);" SRID="<%=oRS("SUBROGATION_DETECTION_RULE_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("SCORE"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("RULE_ID"))%></td>
	</tr>
<%
			oRS.MoveNext
		Loop
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	End If
%>
</tbody>
</table>
</div>
</BODY>
</HTML>


