<%@ Language=VBScript%>
<!--#include file="..\..\lib\genericSQL.asp"-->
<html>
<head>
<META name="VI60_defaultClientScript" content="VBScript">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<link rel="stylesheet" type="text/css" href="../../FNSDESIGN.css">
<title>Agent Billing Report</title>
<SCRIPT LANGUAGE="vbscript">
Sub window_onload
'selMonth.selectedIndex = Month(Date) - 1
'selYear.selectedIndex = 1
Parent.frames("MAIN").location.href = "blank.htm"
End Sub

Sub cmdRun_onclick
dim cRepDate, cAHS, cCustName
'---------------------------------------
' ILOG issue MROU-2726
'Modified By R.Narayan
cRepDateFrom = document.all.txtFrom.value
cRepDateTo = document.all.txtTo.value
'----------------------------------------------
cAHS = "23"
cCustName = "Agent Callflow"

if document.all.selAGTClient.options(document.all.selAGTClient.selectedIndex).value <> "ALL" then
	cAHS = document.all.selAGTClient.options(document.all.selAGTClient.selectedIndex).value
	cCustName = document.all.selAGTClient.options(document.all.selAGTClient.selectedIndex).innerText
end if
'
cmdRun.disabled = true
	
if document.all.Internal.checked then
	Parent.frames("MAIN").location.href = "../Billing/SFR.asp?AHS=" & cAHS & "&DATEFROM=" & cRepDateFrom & "&DATETO=" & cRepDateTo & "&AgtIntBill="
else
	if msgbox("WARNING: This report will Email invoices directly to Customers." & vbcrlf & "Select 'Ok' to proceed, otherwise select 'Cancel'.", 49, "Agent billing") = 1 then
		'Parent.frames("MAIN").location.href = "../Billing/ss.asp?AHS=" & cAHS & "&DATE=" & cRepDate & "&CUSTNAME=" & cCustName	& "&AGT=Y"
		Parent.frames("MAIN").location.href = "verify.asp?AHS=" & cAHS & "&DATEFROM=" & cRepDateFrom & "&DATETO=" & cRepDateTo & "&AGT=Y"
	else
		cmdRun.disabled = false
	end if
end if
End Sub

sub selAGTClient_onchange
if document.all.selAGTClient.options(document.all.selAGTClient.selectedIndex).value = "ALL" then
	document.all.Internal.disabled = false
	document.all.CI.checked = true
else
	document.all.CI.checked = true
	document.all.Internal.checked = false
	document.all.Internal.disabled = true
end if
end sub
</SCRIPT>
</head>
<body bgcolor="Seashell" topmargin="0" leftmargin="0">
<div align="left">
<table border="0" width="66%">
<tr>
<td CLASS="GrpLabel" WIDTH="70" HEIGHT="12"><font face="Verdana, Helvetica,	Arial"><nobr>&nbsp;» Billing Report - Select a Report</font></td>
</tr>
</table>
</div>
<div align="left">
<table border="0" width="80%">
<tr>
<td class="Label" width="7%">Client:<br>
</td>
<td width="24%">
	<select name="selCompany" size="1" class="label">
		<option selected value="AGT">Agent Callflow</option>
	</select></td>
<td class="Label" valign="top" style="width: 12%">
	From:<br>
	<input MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtFrom" ID="Text1">
</td>
<td class="Label" width="11%" valign="top">
	To:<br>
	<input MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="TxtTo" ID="Text2">
</td>
<td class="Label">
	<p></p>
</td>
<td width="29%" align="left">&nbsp; 
	<input id="cmdRun" name="cmdRun" CLASS="StdButton" type="button" value=" Run " width="100">
</td>
</tr>
</table>
<div style="position:absolute; margin-top:7px; margin-left:50px; visibility: visible" class="Label" id="AGTAccount">Agent:</div>
<div style="position:absolute; margin-top:18px;	margin-left:50px; visibility: visible" class="Label" id="AGTDiv">
<table ID="T1" cellspacing="0" cellpadding="0">
<tr>
<td>
	<select name="selAGTClient" size="1" class="label" ID="Select2">
		<option selected value="ALL">All</option>
		<%
		cSQL = "Select name, accnt_hrcy_step_id	from account_hierarchy_step	where parent_node_id=23	And	ACTIVE_STATUS='ACTIVE' order by	name"
		Set	oRS	= Conn.Execute(cSQL)
		do while not oRS.eof
		%>
			<option	value='<%=oRS.Fields("accnt_hrcy_step_id").Value%>'><%=oRS.Fields("name").Value%></option>
		<%
			oRS.moveNext
		loop
		oRS.close
		set	oRS	= nothing
		%>
	</select>
</td>
<td class="Label">
	<p>
	<input type="radio" value="C" checked name="R2" ID="CI">Client Invoice
	<input type="radio" name="R2" value="D" ID="Internal">Internal report</p>
</td>
</tr>
</table>
</div>
</div>
</body>
</html>
<%
Conn.Close()
Set Conn = nothing
%>