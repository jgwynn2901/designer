<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<% Response.Expires = 0 %>
<%
Function ShowDetail(ID)
Select Case ID
	Case "0"
		 ShowDetail = "Fee"
	Case "1"
		 ShowDetail = "Branch Rules"
	Case "2"
		 ShowDetail = "Claim numbers"
	Case "3"
		 ShowDetail = "Branch Assignment"
	Case "4"
		 ShowDetail = "Routing Address"
	Case "5"
		 ShowDetail = "Escalation Rules"
	Case "6"
		 ShowDetail = "Information"
	Case "7"
		 ShowDetail = "Common Routing"
	Case "8"
		 ShowDetail = "Output Definition"
	Case "9"
		 ShowDetail = "Client Routing"
	Case "10"
		 ShowDetail = "Call Flow Rules and Lookups"
	Case "11"
		 ShowDetail = "Attributes"
	Case "12"
		 ShowDetail = "Account Call Flow"
	Case "13"
		 ShowDetail = "EDI UOF Routing"
	Case "14"
		 ShowDetail = "Output OverFlow"
	Case "15"
		 ShowDetail = "Vendor Referral"
	Case "16"
		 ShowDetail = "Fraud Detection"
	Case "17"
		 ShowDetail = "Subrogation"
	Case "20"
		 ShowDetail = "Vendor Eligibility"
End Select
End Function

If Request.QueryString("JOBID") <> "" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
	SQL = ""
	SQL =SQL & "SELECT * FROM MIGRATION_JOB MJ, MIGRATION_DETAIL MD WHERE MJ.JOB_ID=MD.JOB_ID AND MJ.JOB_ID=" & Request.QueryString("JOBID") & " ORDER BY SUBSET_ID"
	Set RS = Conn.Execute(SQL)
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
</HEAD>
<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Job ID</div></td>
			<td class="thd"><div id><nobr>Job Detail ID</div></td>
			<td class="thd"><div id><nobr>Subset ID</div></td>
			<td class="thd"><div id><nobr>ID_TO_MOVE</div></td>
			<td class="thd"><div id><nobr>Subset Start</div></td>
			<td class="thd"><div id><nobr>Subset End</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% 
If Not RS.EOF AND Not RS.BOF Then 
Do While not RS.EOF 
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  JOBID='<%=RS("JOB_DETAIL_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("JOB_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("JOB_DETAIL_ID"))%></td>
	<td NOWRAP CLASS="ResultCell">(<%=renderCell(RS("SUBSET_ID"))%>)  <%= ShowDetail(RS("SUBSET_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ID_TO_MOVE"))%></td>
	<%
	cStart = renderCell(RS("SUBSET_START"))
	cEnd = renderCell(RS("SUBSET_END"))
	if Cint(RS("STATUS_CD")) = 3 then
		if datediff("n", cdate(cStart), cdate(cEnd)) > 20 then
			cStart = "<b>" & cStart & "</b>"
			cEnd = "<b>" & cEnd & "</b>"
		end if
	end if
	%>
	<td NOWRAP CLASS="ResultCell"><%=cStart%></td>
	<td NOWRAP CLASS="ResultCell"><%=cEnd%></td>
	</tr>
<%
RS.MoveNext
Loop
RS.CLose
Else
%>
<tr ID="FieldRow" CLASS="ResultRow" COVID='' >
<td COLSPAN=7 NOWRAP CLASS="ResultCell">No Job selected</td>
</tr>
<% End If %>
</TABLE>
</DIV>
<% End If %>
</BODY>
</HTML>
