<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next

	MCBranchTextLen = 25
	RuleTextLen = 25
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Managed Care Branch Assignment Type Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedMCBARID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedMCBARID = document.all.tblFields.rows(idx).getAttribute("MCBARID")
	Else
		GetSelectedMCBARID = ""
	End If
End Function

function f_LastBARuleRecord
	if document.all.tblFields.Rows.Length <= 2 Then
		f_LastBARuleRecord = true
	else
		f_LastBARuleRecord = false
	end if
end Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" rules=all  cellSpacing="0" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Sequence</div></td>
			<td class="thd"><div id><nobr>State</div></td>			
			<td class="thd"><div id><nobr>FIPS</div></td>			
			<td class="thd"><div id><nobr>Type</div></td>			
			<td class="thd"><div id><nobr>Rule Text</div></td>			
			<td class="thd"><div id><nobr>Branch Office Name</div></td>
			<td class="thd"><div id><nobr>Branch #</div></td>
			<td class="thd"><div id><nobr>Branch ID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	MCBATID = CStr(Request.QueryString("MCBATID"))
	If MCBATID <> "NEW" And MCBATID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT MCBA.MC_BRANCHASSIGNMENTRULE_ID,MCBA.LOB_CD,MCBA.SEQUENCE,MCBA.ROUTING_STATE,MCBA.ROUTING_FIPS,MCBA.MANAGED_CARE_TYPE,R.RULE_TEXT,B.OFFICE_NAME,MCBA.BRANCH_ID, B.BRANCH_NUMBER FROM " &_
				"MC_BRANCHASSIGNMENTRULE MCBA, RULES R, BRANCH B WHERE MCBA.RULE_ID = R.RULE_ID(+) AND " &_
				"MCBA.BRANCH_ID = B.BRANCH_ID(+) AND " &_				
				"MCBA.MC_BRANCHASSIGNMENTTYPE_ID = " & MCBATID & " ORDER BY MCBA.SEQUENCE, MCBA.ROUTING_STATE" 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF And Not RS.BOF
%>


	
	
	<tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);" MCBARID="<%=RS("MC_BRANCHASSIGNMENTRULE_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=RS("LOB_CD")%></td>
	<td NOWRAP CLASS="ResultCell" ><%=RS("SEQUENCE")%></td>
	<td NOWRAP CLASS="ResultCell" ><%=RS("ROUTING_STATE")%></td>
	<td NOWRAP CLASS="ResultCell" ><%=RS("ROUTING_FIPS")%></td>
	<td NOWRAP CLASS="ResultCell" ><%=RS("MANAGED_CARE_TYPE")%></td>
	<td NOWRAP CLASS="ResultCell" TITLE="<%=ReplaceQuotesInText(RS("RULE_TEXT"))%>"><%=TruncateText(RS("RULE_TEXT"),RuleTextLen)%></td>
	<td NOWRAP CLASS="ResultCell" TITLE="<%=RS("OFFICE_NAME")%>"><%=TruncateText(RS("OFFICE_NAME"),MCBranchTextLen)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=RS("BRANCH_NUMBER")%></td>
	<td NOWRAP CLASS="ResultCell" ><%=RS("BRANCH_ID")%></td>
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


