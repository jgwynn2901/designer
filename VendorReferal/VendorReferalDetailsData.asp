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
<title>Vendor Referral Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
function f_LastBARuleRecord
	if document.all.tblFields.Rows.Length <= 2 Then
		f_LastBARuleRecord = true
	else
		f_LastBARuleRecord = false
	end if
end Function
Function GetSelectedBARID
	dim idx
	
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedBARID = document.all.tblFields.rows(idx).getAttribute("BARID")
	Else
		GetSelectedBARID = ""
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
		  
		    <td class="thd"><div id><nobr>Sequence</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>Rule ID</div></td>			
			<td class="thd"><div id><nobr>Rule Text</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	BATID = CStr(Request.QueryString("BATID"))
	If BATID <> "NEW" And BATID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		  SQLST = "SELECT VRR.SEQUENCE,VRR.VENDOR_REFERRAL_RULE_ID,VRR.VENDOR_REFERRAL_TYPE_ID,"
		  SQLST=SQLST & "VRR.DESCRIPTION,VRR.RULE_ID,R.RULE_TEXT" 
          SQLST=SQLST & " FROM VENDOR_REFERRAL_RULE VRR,VENDOR_REFERRAL_TYPE VRT,RULES R"
          SQLST=SQLST & " WHERE VRR.RULE_ID = R.RULE_ID(+)" 
          SQLST=SQLST & "AND VRR.VENDOR_REFERRAL_TYPE_ID=  VRT.VENDOR_REFERRAL_TYPE_ID"
          SQLST=SQLST & " AND VRR.VENDOR_REFERRAL_TYPE_ID = "& BATID 
		  SQLST=SQLST & " ORDER BY VRR.SEQUENCE" 
				
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
		
		
%>
    <tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);" BARID="<%=RS("VENDOR_REFERRAL_RULE_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("SEQUENCE"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("DESCRIPTION"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("RULE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" TITLE="<%=ReplaceQuotesInText(RS("RULE_TEXT"))%>"><%=renderCell(TruncateText(RS("RULE_TEXT"),RuleTextLen))%></td>
	
	
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


