<%@ Language=VBScript%>
<!--#include file="..\..\lib\genericSQL.asp"-->
<!--#include file="billing.inc"-->
<%

dim lChkDates, lIsOld

lChkDates = InStr(1, Request.QueryString, "CHKDATE", 1) <> 0
if lChkDates then
	doCheck
end if
if InStr(1, Request.QueryString, "AgtIntBill", 1) <> 0 then
	if not Application("lExecutingBillingReport") then
		Response.redirect "runReport.asp?" & Request.QueryString
	else
		Response.redirect "inUse.htm"
	end if
end if

sub doCheck
dim cAHS, cStartDate, oConn, oRS, nResult, cSQL
dim cCustName, cCustCode

'	ra		- the monthly report has been replaced by date ranges, so this is not needed
'cSQL = "Select * from BILLING_HISTORY Where MMM_YYYY='" & UCase(cStartDate) & "' and AHS_ID='" & cAHS & "'"
'Set oRS = Conn.Execute(cSQL)
'lIsOld = not oRS.eof
'oRS.close
lIsOld = false
set oRS = nothing
if not lIsOld then
	if not Application("lExecutingBillingReport") then
		Response.redirect "runReport.asp?" & Request.QueryString
	else
		Response.redirect "inUse.htm"
	end if
else
	Response.redirect "NewOrExisting.asp?" & Request.QueryString
end if
end sub

'*************CISG client*********8
     	   
	Dim isCISG
	isCISG = false
	if left(getInstanceName,4) = "CISG" then
		isCISG = true
	End If

'******************************************8

'*************AME client*********8

	Dim isAME
	isAME = false
	if left(getInstanceName,3) = "AME" then
		isAME = true
	End If

'******************************************8

'*************SEL client********* PMAC-1839 **8 
	
	Dim isSEL
	isSEL = false
	if left(getInstanceName,3) = "SEL" then
		isSEL = true
	End If

'******************************************8

'*************TOW client********* TPAL-0146 **8 

    Dim isTOW
    isTOW = false
    if left(getInstanceName,3) = "TOW" then
    isTOW = true
    End If

'******************************************8


%>
<html>
<head>
<META name="VI60_defaultClientScript" content="VBScript">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="..\..\FNSDESIGN.css">
<title>Service Fee Reports</title>
<SCRIPT LANGUAGE="vbscript">
Sub window_onload
'selMonth.selectedIndex = Month(Date) - 1
'selYear.selectedIndex = 1
Parent.frames("MAIN").location.href = "blank.htm"
End Sub

Sub cmdRun_onclick
dim cRepDateFrom, cRepDateTo, cAHS, cCustName, lIsCCE, lIsAgent
cMsg = validateDates
if len(cMsg) <> 0 then
	msgbox cMsg
	exit sub
end if
cRepDateFrom = document.all.txtFrom.value
cRepDateTo = document.all.txtTo.value

	Select case selCompany.value
		case "AIG"
			cAHS = <%=AIGNo%>
			cCustName = "<%=AIGName%>"
		case "GBS"
			cAHS = <%=GBSNo%>
			cCustName = "<%=GBSName%>"
		case "MAR"
			cAHS = <%=MARNo%>
			cCustName = "<%=MARName%>"
		case "FRE"
			cAHS = <%=FRENo%>
			cCustName = "<%=FREName%>"
		case "FMT"
			cAHS = <%=FMTNo%>
			cCustName = "<%=FMTName%>"
		case "ULI"
			cAHS = <%=ULINo%>
			cCustName = "<%=ULIName%>"
		case "TIG"
			cAHS = <%=TIGNo%>
			cCustName = "<%=TIGName%>"
		case "MCD"
			cAHS = <%=MCDNo%>
			cCustName = "<%=MCDName%>"
		case "KMP"
			cAHS = <%=KMPNo%>
			cCustName = "<%=KMPName%>"
	    case "KMPC"
			cAHS = <%=KMPCATNo%>
			cCustName = "<%=KMPCATName%>"
		case "CCE"
			cAHS = <%=CCENo%>
			cCustName = "<%=CCEName%>"
		case "LAC"
			cAHS = <%=LACNo%>
			cCustName = "<%=LACName%>"
		case "WIG"
			cAHS = <%=WIGNo%>
			cCustName = "<%=WIGName%>"
		case "MGC"
			cAHS = <%=MGCNoIRC%>
			cCustName = "<%=MGCNameIRC%>"
		case "MGCR"
			cAHS = <%=MGCNoReg%>
			cCustName = "<%=MGCNameReg%>"
	    case "PCN"
			cAHS = <%=PCNNo%>
			cCustName = "<%=PCNName%>"
		case "ARG"
			cAHS = <%=ARGNo%>
			cCustName = "<%=ARGName%>"
		case "AIM"
			cAHS = <%=AIMNo%>
			cCustName = "<%=AIMName%>"
		case "NBIC"
			cAHS = <%=NBICNo%>
			cCustName = "<%=NBICName%>"
		 case "ONB"
			cAHS = <%=ONBNo%>
			cCustName = "<%=ONBName%>"
	   case "CSG"
			cAHS = <%=CSGNoCall%>
			cCustName = "<%=CSGNameCall%>"
		case "CSGL"
			cAHS = <%=CSGNoOnline%>
			cCustName = "<%=CSGNameOnline%>"
	   case "CSGSUM"
			cAHS = <%=CSGNoSummary%>
			cCustName = "<%=CSGNameSummary%>"
		case "FG"
			cAHS = <%=FGNo%>
			cCustName = "<%=FGName%>"
		case "CSAA"
			cAHS = <%=CSAANo%>
			cCustName = "<%=CSAAName%>"
		case "AAA"
			cAHS = <%=AAANo%>
			cCustName = "<%=AAAName%>"
		case "STA"
			cAHS = <%=STANo%>
			cCustName = "<%=STAName%>"
		case "RDC"
			cAHS = <%=RDCNo%>
			cCustName = "<%=RDCName%>"
		case "AIK"
			cAHS = <%=AIKNo%>
			cCustName = "<%=AIKName%>"
		case "CHB"
	        cAHS = <%=CHBNo%>
	        cCustName = "<%=CHBName%>"
	  case "WMA"
			 cAHS = <%=WMANo%>
			 cCustName = "<%=WMAName%>"
		case "SHPR"
			 cAHS = <%=SHPRNo%>
			 cCustName = "<%=SHPRName%>"
	  case "CVG"
			 cAHS = <%=CVGNo%>
			 cCustName = "<%=CVGName%>"
		case "CIR"
			 cAHS = <%=CIRNo%>
			 cCustName = "<%=CIRName%>"
		case "CRWASP"
			 cAHS = <%=CRWNoASP%>
			 cCustName = "<%=CRWNameASP%>"
		case "CRWFNS"
			 cAHS = <%=CRWNoFNS%>
			 cCustName = "<%=CRWNameFNS%>"
		case "MTS"
			 cAHS = <%=MTSNo%>
			 cCustName = "<%=MTSName%>"
	   case "MER"
			 cAHS = <%=MERNo%>
			 cCustName = "<%=MERName%>"

		case "BRK"
			 cAHS = <%=BRKNo%>
			 cCustName = "<%=BRKName%>"
		case "AMC"
			 cAHS = <%=AMCNo%>
			 cCustName = "<%=AMCName%>"
	    case "SEN"
			 cAHS = <%=SENNo%>
			 cCustName = "<%=SENName%>"
	    case "SEA"
			 cAHS = <%=SEANo%>
			 cCustName = "<%=SEAName%>"
		case "ACE"
			 cAHS = <%=ACENo%>
			 cCustName = "<%=ACEName%>"
	    case "EMC"
			cAHS = <%=EMCNo%>
			cCustName = "<%=EMCName%>"
		'KFAB-6227
		case "AFFM"
			cAHS = <%=AFFMNo%>
			cCustName = "<%=AFFMName%>"
		case "NTW"
			cAHS = <%=NTWNo%>
			cCustName = "<%=NTWName%>"
		case "Canal"
			cAHS = <%=CNLNo%>
			cCustName = "<%=CNLName%>"
		'---To Be Removed
		'case "Canal3-in-1"
		'	cAHS = <%=CNLNewNo%>
		'	cCustName = "<%=CNLNameNew%>"
		'---To Be Removed
		case "RTW"
			cAHS = <%=RTWNo%>
			cCustName = "<%=RTWName%>"
		'Added below for HML client on 30-Sep-05
		case "HML"
			cAHS = <%=HMLNo%>
			cCustName = "<%=HMLName%>"
        case "ALM"
			cAHS = <%=ALMNo%>
			cCustName = "<%=ALMName%>"
		case "SEL"
			cAHS = <%=SELNo%>
			cCustName = "<%=SELName%>"
		case "SRS"
			cAHS = <%=SRSNo%>
			cCustName = "<%=SRSName%>"
		case "UNI"
			cAHS = <%=UNINo%>
			cCustName = "<%=UNIName%>"
		case "PMCO"
			cAHS = <%=PMCONo%>
			cCustName = "<%=PMCOName%>"
		case "TGC"
			cAHS = <%=TGCNo%>
			cCustName = "<%=TGCName%>"
		'TPAL-0146
		case "TGCASP"
			cAHS = <%=TGCNoASP%>
			cCustName = "<%=TGCNameASP%>"
		case "ESIS"
			cAHS = <%=ESISNo%>
			cCustName = "<%=ESISName%>"
		case "EVR"
			cAHS = <%=EVRNo%>
			cCustName = "<%=EVRName%>"
		case "SAF"
			cAHS = <%=SAFNo%>
			cCustName = "<%=SAFName%>"
		case "ANI"
			cAHS = <%=ANINo%>
			cCustName = "<%=ANIName%>"
		case "AME"
			cAHS = <%=AMENo%>
			cCustName = "<%=AMEName%>"
	end select

	if cAHS = 11 then	'	CCE
		lIsCCE = true
		if document.all.selCCEAccount.options(document.all.selCCEAccount.selectedIndex).value <> "ALL" then
			cAHS = document.all.selCCEAccount.options(document.all.selCCEAccount.selectedIndex).value
			cCustName = document.all.selCCEAccount.options(document.all.selCCEAccount.selectedIndex).innerText
		end if
	end if
	'
	cmdRun.disabled = true

	if document.all.Detail.checked then
		checkReportDates cRepDateFrom, cRepDateTo, cAHS, selCompany.value, cCustName
	else
		if lIsCCE then
			Parent.frames("MAIN").location.href = "ss.asp?AHS=" & cAHS & "&DATEFROM=" & cRepDateFrom & "&DATETO=" & cRepDateTo & "&CUSTNAME=" & cCustName	& "&CCE=Y"
	   else
	        Parent.frames("MAIN").location.href = "ss.asp?AHS=" & cAHS & "&DATEFROM=" & cRepDateFrom & "&DATETO=" & cRepDateTo & "&CUSTNAME=" & cCustName
		end if
	end if
End Sub

function validateDates
validateDates = ""
if len(document.all.txtFrom.value) = 0 then
	validateDates = "Please enter a From date"
elseif len(document.all.txtTo.value) = 0 then
	validateDates = "Please enter a To date"
elseif not isDate(document.all.txtFrom.value) then
	validateDates = "The From date has an incorrect format"
elseif not isDate(document.all.txtTo.value) then
	validateDates = "The To date has an incorrect format"
elseif CDate(document.all.txtTo.value) < CDate(document.all.txtFrom.value) then
	validateDates = "Incorrect date range."
end if
end function

sub cmdreset_onClick

parent.frames("MAIN").location.href = "reset.asp"

end sub
sub checkReportDates(cRepDateFrom, cRepDateTo, cAHS, cCustCode, cCustName)
Parent.frames("MAIN").location.href = "SFR.asp?AHS=" & cAHS & "&DATEFROM=" & cRepDateFrom & "&DATETO=" & cRepDateTo & "&CHKDATE=&CUSTCODE=" & cCustCode  & "&CUSTNAME=" & cCustName
end sub

sub runReport(cStartDate, cAHS)
Parent.frames("MAIN").location.href = "runReport.asp?AHS=" & cAHS & "&DATE=" & cStartDate
end sub

sub selCompany_onchange
dim cSelectedClient

cSelectedClient = document.all.selCompany.options(document.all.selCompany.selectedIndex).value
 document.all.Summary.disabled = false
 document.all.cmdrun.disabled = false
document.all.CCEAccount.style.visibility = "hidden"
document.all.selCCEAccount.style.visibility = "hidden"
if cSelectedClient = "CCE" then
	document.all.CCEAccount.style.visibility = "visible"
	document.all.selCCEAccount.style.visibility = "visible"
	selCCEAccount_onchange
	parent.document.all.F1.rows = "100,*"
elseif cSelectedClient = "AGT" then
	document.all.CCEAccount.style.visibility = "hidden"
	document.all.selCCEAccount.style.visibility = "hidden"
	document.all.Summary.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false
	document.all.Detail.disabled = true
	parent.document.all.F1.rows = "110,*"
elseif cSelectedClient = "AIK" then
	document.all.Summary.disabled = true
	document.all.Detail.checked = true
elseif cSelectedClient = "ONB" then
	document.all.Summary.disabled = true
	document.all.Detail.checked = true
elseif  cSelectedClient = "KMPC" then
	document.all.Summary.checked = true
	document.all.Detail.disabled = true
 elseif  cSelectedClient = "CVG" then
	document.all.Summary.disabled = true
	document.all.Detail.checked = true
elseif cSelectedClient = "MTS" then
	document.all.Summary.disabled = true
	document.all.Detail.checked = true
 elseif cSelectedClient = "WMAOH" then
	document.all.Detail.disabled = true
	document.all.Summary.checked = true
  elseif cSelectedClient = "MER" then
	document.all.Detail.disabled = true
	document.all.Summary.checked = true
 elseif cSelectedClient = "BRK" then
	document.all.Detail.disabled = true
	document.all.Summary.checked = true
 elseif  cSelectedClient = "CSG"  THEN
    document.all.Summary.disabled = true
    document.all.Detail.checked = true
elseif cSelectedClient = "CSGL"  THEN
     document.all.Summary.disabled = true
     document.all.Detail.checked = true
elseif  cSelectedClient = "CSGSUM" then
	document.all.Detail.disabled = true
	document.all.Summary.checked = true
elseif cSelectedClient = "SEN" then
	document.all.Detail.checked = true
	document.all.Summary.disabled = true
elseif cSelectedClient = "ACE" then
	document.all.Detail.checked = true
	document.all.Summary.disabled = true
elseif cSelectedClient = "CSAA" then
	document.all.Summary.disabled = false
	document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false
elseif cSelectedClient = "AAA" then
	document.all.Summary.disabled = false
	document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false
elseif cSelectedClient = "NBIC" then
	document.all.Summary.disabled = false
	document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false
elseif cSelectedClient = "EMC" then
	document.all.Detail.checked = true
	'AMCF-0163 07/27/2005
	'document.all.Summary.disabled = true
	document.all.Summary.disabled = false
'KFAB-6227
elseif cSelectedClient = "AFFM" then
	document.all.Detail.checked = true
	document.all.Summary.disabled = false
elseif cSelectedClient = "NTW" then
	document.all.Detail.checked = true
	document.all.Summary.disabled = true
elseif cSelectedClient = "Canal" then
	document.all.Summary.checked = true
	document.all.Detail.disabled = false
'elseif cSelectedClient = "Canal3-in-1" then
'	document.all.Summary.checked = true
'	document.all.Detail.disabled = false
elseif cSelectedClient = "RTW" then
	document.all.Summary.checked = true
	document.all.Detail.disabled = false
elseif cSelectedClient = "HML" then
	document.all.Summary.checked = true
	document.all.Detail.disabled = false
elseif cSelectedClient = "ALM" then
	document.all.Summary.checked = true
	document.all.Detail.disabled = false
elseif cSelectedClient = "SEL" then
    document.all.Summary.disabled = true
	document.all.Detail.checked = true
elseif cSelectedClient = "SRS" then
    document.all.Summary.checked = true
	document.all.Detail.checked = true
elseif cSelectedClient = "UNI" then
    document.all.Summary.disabled = true
	document.all.Detail.checked = true
elseif cSelectedClient = "PMCO" then
	document.all.Summary.disabled = false
	document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false

elseif cSelectedClient = "TGC" or cSelectedClient = "TGCASP" then
	document.all.Summary.disabled = false
	document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false
elseif cSelectedClient = "ESIS" then
    document.all.Detail.disabled = false
	document.all.Summary.checked = true
elseif cSelectedClient = "EVR" then
    document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Summary.disabled = false
	document.all.Detail.checked = false
elseif cSelectedClient = "SAF" then
    document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Summary.disabled = false
	document.all.Detail.checked = false
elseif cSelectedClient = "ANI" then
	document.all.Summary.disabled = true
	document.all.Detail.checked = true
elseif cSelectedClient = "AME" then
	document.all.Summary.disabled = false
	document.all.Detail.disabled = false
	document.all.Summary.checked = true
	document.all.Detail.checked = false
else
	parent.document.all.F1.rows = "70,*"
end if
end sub

sub selCCEAccount_onchange
if document.all.selCCEAccount.options(document.all.selCCEAccount.selectedIndex).value <> "ALL" then
	document.all.Summary.disabled = false
	document.all.Summary.checked = true
	if document.all.selCCEAccount.options(document.all.selCCEAccount.selectedIndex).value = "21675233" then
		document.all.Detail.disabled = false
	else
		document.all.Detail.disabled = true
	end if
	document.all.Detail.checked = false
else
	document.all.Summary.disabled = true
	document.all.Summary.checked = false
	document.all.Detail.checked = true
	document.all.Detail.disabled = false
end if
end sub

</SCRIPT>
</head>
<body bgcolor="Seashell" topmargin="0" leftmargin="0">
<div align="left">
<table border="0" width="66%">
<tr>
<td CLASS="GrpLabel" WIDTH="70" HEIGHT="12"><font face="Verdana, Helvetica, Arial"><nobr>&nbsp;» Billing Report - Select a Report</font></td>
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
   <% if isCISG = true then %>
		 <option value="CSG"><%=CSGNameCall%></option>
		 <option value="CSGL"><%=CSGNameOnline%></option>
	<%  elseif isAME = true then %>
		 <option value="AME"><%=AMEName%></option>	  
			<!-- PMAC-1839 -->
	<%  elseif isSEL = true then %>        
		 <option value="SEL"><%=SELName%></option>
	<!-- TPAL-0146 -->
	<%  elseif isTOW = true then %>
	    <option value="TGCASP"><%=TGCNameASP%></option>
	<%  else	%>
	     <option value="CSGSUM"><%=CSGNameSummary%></option>
		 <option value="AAA"><%=AAAName%></option>
	     <option value="ACE"><%=ACEName%></option>
	     <!--KFAB-6227-->
	     <option value="AFFM"><%=AFFMName%></option>
		 <option selected value="AIG"><%=AIGName%></option>
		  <option value="AIK"><%=AIKName%></option>
		  <option value="AIM"><%=AIMName%></option>
  		  <option value="ANI"><%=ANIName%></option>
		  <option value="ALM"><%=ALMName%></option>
		  <option value="AMC"><%=AMCName%></option>
		  <option value="ARG"><%=ARGName%></option>
		  <option value="BRK"><%=BRKName%></option>
		  <option value="Canal"><%=CNLName%></option>
		  <!--option value="Canal3-in-1"><%=CNLNameNew%></option-->
		  <option value="CHB"><%=CHBName%></option>
		  <option value="CCE"><%=CCEName%></option>
		  <option value="CIR"><%=CIRName%></option>
		  <option value="CRWASP"><%=CRWNameASP%></option>
		  <option value="CRWFNS"><%=CRWNameFNS%></option>
		  <option value="CVG"><%=CVGName%></option>
		  <option value="CSAA"><%=CSAAName%></option>
		  <option value="EMC"><%=EMCName%></option>
		  <option value="FG"><%=FGName%></option>
		  <option value="FMT"><%=FMTName%></option>
		  <option value="FRE"><%=FREName%></option>
		  <option value="GBS"><%=GBSName%></option>
		  <option value="HML"><%=HMLName%></option>
		  <option value="KMP"><%=KMPName%></option>
		  <option value="KMPC"><%=KMPCATName%></option>
		  <option value="LAC"><%=LACName%></option>
		  <option value="MAR"><%=MARName%></option>
		  <option value="MCD"><%=MCDName%></option>
		  <option value="MER"><%=MERName%></option>
		  <option value="MGC"><%=MGCNameIRC%></option>
		  <option value="MGCR"><%=MGCNameReg%></option>
		  <option value="MTS"><%=MTSName%></option>
		  <option value="NTW"><%=NTWName%></option>
  		  <option value="NBIC"><%=NBICName%></option>
		  <option value="ONB"><%=ONBName%></option>
		  <option value="PCN"><%=PCNName%></option>
		  <option value="RDC"><%=RDCName%></option>
		  <option value="RTW"><%=RTWName%></option>
		   <option value="SEA"><%=SEAName%></option>
		   <option value="SEL"><%=SELName%></option>
		   <option value="SRS"><%=SRSName%></option>
		   <option value="SEN"><%=SENName%></option>
		  <option value="SHPR"><%=SHPRName%></option>
		  <option value="STA"><%=STAName%></option>
		  <option value="TIG"><%=TIGName%></option>
		  <option value="ULI"><%=ULIName%></option>
		  <option value="WIG"><%=WIGName%></option>
		  <option value="WMA"><%=WMAName%></option>
		  <option value="UNI"><%=UNIName%></option>
		  <option value="PMCO"><%=PMCOName%></option>
		  <option value="ESIS"><%=ESISName%></option>
		  <option value="EVR"><%=EVRName%></option>
		  <option value="TGC"><%=TGCName%></option>
		  <option value="SAF"><%=SAFName%></option>
	<% end if %>

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
<% if isCISG = true then %>
<p><input type="radio" value="S"  ID="Summary">Summary<input type="radio" name="R1" value="D" checked name="R1" ID="Detail">Detail</p>
<%elseif isSEL = true then %>
<p><input type="radio" value="S"  ID="Summary" disabled >Summary<input type="radio" name="R1" value="D" checked name="R1" ID="Detail">Detail</p>
<% else %>
	<p><input type="radio" value="S" checked name="R1" ID="Summary">Summary<input type="radio" name="R1" value="D" ID="Detail">Detail</p>
 <%end if%>
</td>
<td width="29%" align="left">&nbsp;
	<input id="cmdRun" name="cmdRun" CLASS="StdButton" type="button" value= "Run" width="100">
	<input id="cmdReset" name="cmdReset" CLASS="StdButton" type="button" value= "Reset" width="100">
</td>
</tr>
</table>
<div style="position:absolute; margin-top:7px; margin-left:50px; visibility: hidden" class="Label" id="CCEAccount">CCE Account:</div>


<div style="position:absolute; margin-top:20px; margin-left:50px; visibility: hidden" class="Label" id="Div1">
	<select name="selCCEAccount" size="1" class="label" ID="SelectCCEAccnt" style="visibility: hidden">
		<option selected value="ALL">All</option>
		<%
		cSQL = "Select name, accnt_hrcy_step_id from account_hierarchy_step where parent_node_id=11 And ACTIVE_STATUS='ACTIVE'AND accnt_hrcy_step_id <> 10786278 order by name"
		Set oRS = Conn.Execute(cSQL)
		do while not oRS.eof
		%>
			<option value='<%=oRS.Fields("accnt_hrcy_step_id").Value%>'><%=oRS.Fields("name").Value%></option>
		<%
			oRS.moveNext
		loop
		oRS.close
		set oRS = nothing
		%>
	</select>
</div>
</body>
</html>
<%
Conn.Close()
Set Conn = nothing
%>
