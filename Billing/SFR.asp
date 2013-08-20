<%@ Language=VBScript%>
<!--#include file="..\lib\genericSQL.asp"-->
<%
const Fre = "Fremont"
const Mar = "Marriot"
const GBS = "Gallagher Basset Services, Inc"
const AIG = "American International Group"

dim lChkDates, lIsOld

lChkDates = InStr(1, Request.QueryString, "CHKDATE", 1) <> 0
if lChkDates then
	doCheck
end if	

sub doCheck
dim cAHS, cStartDate, oRS, nResult, cSQL
dim cCustName, cCustCode

cStartDate = Request.QueryString("DATE")
cAHS = Request.QueryString("AHS")
cCustName = Request.QueryString("CUSTNAME")
cCustCode = Request.QueryString("CUSTCODE")

cSQL = "Select * from BILLING_HISTORY Where MMM_YYYY='" & UCase(cStartDate) & "' and AHS_ID='" & cAHS & "'"
Set oRS = Conn.Execute(cSQL)
lIsOld = not oRS.eof
oRS.close
set oRS = nothing
if not lIsOld then
	Response.redirect "runReport.asp?" & Request.QueryString
else
	Response.redirect "NewOrExisting.asp?" & Request.QueryString
end if
end sub
%>
<html>

<head>
<META name=VI60_defaultClientScript content=VBScript>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="..\..\FNSDESIGN.css">
<title>Service Fee Reports</title>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
selMonth.selectedIndex = Month(Date) - 1
selYear.selectedIndex = 1
End Sub

Sub cmdRun_onclick
dim cRepDate, cAHS, cCustName

	cRepDate = selMonth.value & selYear.value 
	Select case selCompany.value 
		case "AIG"
			cAHS = 62
			cCustName = "<%=AIG%>"
		case "GBS"
			cAHS = 46
			cCustName = "<%=GBS%>"
		case "MAR"
			cAHS = 72
			cCustName = "<%=Mar%>"
		case "FRE"
			cAHS = 75
			cCustName = "<%=Fre%>"
	end select
	cmdRun.disabled = true
	checkReportDates cRepDate, cAHS, selCompany.value, cCustName
End Sub

sub checkReportDates(cStartDate, cAHS, cCustCode, cCustName)
Parent.frames("MAIN").location.href = "SFR.asp?AHS=" & cAHS & "&DATE=" & cStartDate & "&CHKDATE=&CUSTCODE=" & cCustCode  & "&CUSTNAME=" & cCustName
end sub

sub runReport(cStartDate, cAHS)
Parent.frames("MAIN").location.href = "runReport.asp?AHS=" & cAHS & "&DATE=" & cStartDate
end sub
-->
</SCRIPT>
</head>

<body bgcolor="Seashell" topmargin="0" leftmargin="0">
<div align="left">
<table border="0" width="63%">
  <tr>
	<td CLASS="GrpLabel" WIDTH="50" HEIGHT="12"><font face="Verdana, Helvetica, Arial"><nobr>&nbsp;» Billing Report - Select a Report</font></td>
  </tr>
</table>
</div><div align="left">

<table border="0" width="80%">
  <tr>
	<td class="Label" width="7%">Account:<br>	
	</td>
    <td width="24%"><select name="selCompany" size="1" class="label">
      <option selected value="AIG"><%=AIG%></option>
      <option value="FRE"><%=Fre%></option>
      <option value="GBS"><%=GBS%></option>
      <option value="MAR"><%=Mar%></option>
    </select></td>
	<td class="Label" width="6%">Period:<br>	
	</td>
    <td width="8%"><select name="selMonth" size="1" class="label">
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
    </select></td>
    <td width="11%"><select name="selYear" size="1" class="label">
      <option value="<%=Year(Date) - 1%>"><%=Year(Date) - 1%></option>
      <option value="<%=Year(Date)%>"><%=Year(Date)%></option>
      <option value="<%=Year(Date) + 1%>"><%=Year(Date) + 1%></option>
    </select></td>
    <td width="29%"  align ="left">&nbsp;
	<input id="cmdRun" name="cmdRun" CLASS="StdButton" type="button" value="  Run  " width="100">
  </tr>
  <tr>
  <span style="font-family: Verdana, Helvetica, Arial; font-size: 11pt" ID=spanData></span>
  </tr>
</table>
</div>
</body>
</html>
