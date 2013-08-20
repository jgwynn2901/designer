<%@ Language=VBScript%>
<%
dim oConn, oRS

set oConn = server.CreateObject("ADODB.Connection")
oConn.Open "DSN=HALP;UID=FNSOWNER;PWD=CTOWN_PROD"

%>
<html>
<head>
<META name="VI60_defaultClientScript" content="VBScript">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="..\..\FNSDESIGN.css">
<title>Health Alliance Billing Report</title>
<SCRIPT LANGUAGE="vbscript">
Sub window_onload
selMonth.selectedIndex = Month(Date) - 1
selYear.selectedIndex = 1
Parent.frames("MAIN").location.href = "blank.htm"
End Sub

Sub cmdRun_onclick
dim cRepDate, cAHS, cCustName, cReport

cRepDate = selMonth.value & selYear.value 

cAHS = document.all.selCompany.options(document.all.selCompany.selectedIndex).value
cCustName = document.all.selCompany.options(document.all.selCompany.selectedIndex).innerText
'
cmdRun.disabled = true
	
if document.all.OE.checked then
	cReport = "OE"
elseif document.all.IE.checked then
	cReport = "IE"
elseif document.all.OR.checked then
	cReport = "OR"
elseif document.all.IR.checked then
	cReport = "IR"
elseif document.all.IC.checked then
	cReport = "IC"
elseif document.all.FD.checked then
	cReport = "FD"
end if
Parent.frames("MAIN").location.href = "../Billing/ss.asp?AHS=" & cAHS & "&cust=" & cCustName & "&report=" & cReport & "&date=" & cRepDate & "&HABill="
End Sub
</SCRIPT>
</head>
	<body bgcolor="Seashell" topmargin="0" leftmargin="0">
		<div align="left">
			<table border="0" width="66%">
				<tr>
					<td CLASS="GrpLabel" WIDTH="70" HEIGHT="12"><font face="Verdana, Helvetica, Arial"><nobr>&nbsp;» 
							Health Alliance Billing Report - Select a State</font></td>
				</tr>
			</table>
		</div>
		<div align="left">
			<table border="0" width="80%">
				<tr>
					<td class="Label" width="7%">State:<br>
					</td>
					<td width="24%">
						<select name="selCompany" size="1" class="label">
							<%
					cSQL = "Select name, accnt_hrcy_step_id from account_hierarchy_step where parent_node_id=1 And ACTIVE_STATUS='ACTIVE' order by name"
					Set oRS = oConn.Execute(cSQL)
					do while not oRS.eof
					%>
							<option value='<%=oRS.Fields("accnt_hrcy_step_id").Value%>'><%=oRS.Fields("name").Value%></option>
							<%
						oRS.moveNext
					loop
					oRS.close
					set oRS = nothing
					%>
						</select></td>
					<td class="Label" width="6%">Period:<br>
					</td>
					<td width="8%">
						<select name="selMonth" size="1" class="label">
							<option value="Jan">Jan</option>
							<option value="Feb">Feb</option>
							<option value="Mar">Mar</option>
							<option value="Apr">Apr</option>
							<option value="May">May</option>
							<option value="Jun">Jun</option>
							<option value="Jul">Jul</option>
							<option value="Aug">Aug</option>
							<option value="Sep">Sep</option>
							<option value="Oct">Oct</option>
							<option value="Nov">Nov</option>
							<option value="Dec">Dec</option>
						</select>
					</td>
					<td width="11%">
						<select name="selYear" size="1" class="label">
							<option value="<%=Year(Date) - 1%>"><%=Year(Date) - 1%></option>
							<option value="<%=Year(Date)%>"><%=Year(Date)%></option>
							<option value="<%=Year(Date) + 1%>"><%=Year(Date) + 1%></option>
						</select>
					</td>
					<td class="Label">
						<p></p>
					</td>
					<td width="29%" align="left">&nbsp; <input id="cmdRun" name="cmdRun" CLASS="StdButton" type="button" value="  Run  " width="100">
					</td>
				</tr>
			</table>
			<div style="position:absolute; margin-top:18px; margin-left:10px; visibility: visible" class="Label" id="AGTDiv">
				<table ID="T1" cellspacing="0" cellpadding="0">
					<tr>
						<td class="Label">
							<input type="radio" value="OE" checked name="R2" ID="OE">Outbound Enrollment
						</td>
						<td class="Label">
							<input type="radio" name="R2" value="IE" ID="IE">Inbound Enrollment
						</td>
						<td class="Label">
							<input type="radio" value="OR" name="R2" ID="OR">Outbound Refills
						</td>
						<td class="Label">
							<input type="radio" value="IR" name="R2" ID="IR">Inbound Refills
						</td>
						<td class="Label">
							<input type="radio" value="IC" name="R2" ID="IC">Info Calls
						</td>
						<td class="Label">
							<input type="radio" value="FD" name="R2" ID="FD">Faxes to Doctors
						</td>
					</tr>
				</table>
			</div>
	</body>
</html>
<%
oConn.Close()
Set oConn = nothing
%>
