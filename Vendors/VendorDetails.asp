<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	Response.Expires=0 %>
<!--#include file="..\lib\ZIP.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Branch Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
var g_StatusInfoAvailable = false;
var lServiceTypeChangeDetected = false;
var lServiceDaysChangeDetected = false;

function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
sub fillZip1(cVarName)
dim cURL

oZip.Zip = ""
oZip.State = ""
oZip.City = ""
cURL = "../LIB/ZIPLookupIFrame.asp?ZIP=" & document.all.item(cVarName & "ZIP").value
showModalDialog  cURL, oZip ,"center"
document.all.item(cVarName & "Zip").value = oZip.Zip
document.all.item(cVarName & "State").value = oZip.State
document.all.item(cVarName & "City").value = oZip.City
end sub

sub Txt0Zip_onblur
	doZIP("TXT0ZIP")
end sub

sub Txt1Zip_onblur
	doZIP("TXT1ZIP")
end sub

sub doZIP(cVarName)
dim x, lFocusSet

fillZIP1 left(cVarName, 4)	'	pass only TXT0, TXT1 and so on
on error resume next
for x=0 to document.all.length-1
	if uCase(Trim(document.all(x).name)) = cVarName then
		if err.number = 0 then
			exit for
		elseif err.number <> 438 then	'	no such property
			msgbox "Internal Error: " & err.number & " - " & err.description
			exit for
		else
			err.clear
		end if
	end if
next
lFocusSet = false
for z=x+1 to document.all.length-1
	document.all(z).focus
	if err.number = 0 then
		lFocusSet = true
		exit for
	else
		err.clear
	end if
next
on error goto 0
end sub

function getServiceTypes
	dim i, cResult
	
	cResult = ""
	for i=1 to tblFields.rows.length - 1
		if tblFields.rows(i).cells(2).children(0).checked then
			if len(cResult) <> 0 then
				cResult = cResult & ";"
			end if 
			cResult = cResult & tblFields.rows(i).getAttribute("STID")
		end if
	next
	getServiceTypes = cResult
end function

function getServiceDays
	dim i, cResult
	
	cResult = ""
	for i=1 to 7
		if ServiceDays.rows(1).cells(i).children(0).checked then
			if len(cResult) <> 0 then
				cResult = cResult & "~"
			end if 
			cResult = cResult & UCase(ServiceDays.rows(0).cells(i).innerText) & "$"
			cResult = cResult & ServiceDays.rows(2).cells(i).children(0).value & "$"
			cResult = cResult & ServiceDays.rows(3).cells(i).children(0).value
		end if
	next
	getServiceDays = cResult
end function

Sub PostTo(strURL)
	FrmDetails.action = "VendorSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateVID(inVID)
	document.all.VID.value = inVID
	document.all.spanVID.innerText = inVID
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Function GetVID
	if document.all.VID.value <> "NEW" then
		GetVID = document.all.VID.value
	else
		GetVID = ""
	end if 
End Function

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	dim cErrMsg

	cErrMsg = ""
	ValidateScreenData = true
	If document.all.TxtName.value = "" then
		cErrMsg = "Vendor name is a required field." & vbCRLF
	end if
	If document.all.TxtAddress1.value = "" then
		cErrMsg = cErrMsg & "Address 1 is a required field." & vbCRLF
	end if
	if document.all.Txt0ZIP.value = "" then
		cErrMsg = cErrMsg & "ZIP code is a required field." & vbCRLF
	end if
	if document.all.TxtPhone.value = "" then
		cErrMsg = cErrMsg & "Day time phone is a required field." & vbCRLF
	end if
	if len(getServiceTypes) = 0 then
		cErrMsg = cErrMsg & "Please choose at least one Service provided by this Vendor." & vbCRLF
	end if
	if len(getServiceDays) = 0 then
		cErrMsg = cErrMsg & "Please indicate on which days this Vendor works." & vbCRLF
	end if
	if len(cErrMsg) <> 0 then
		MsgBox cErrMsg, 0, "FNSDesigner"
		ValidateScreenData = false
	end if
End Function

Function ExeSave
	sResult = ""
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.VID.value = "" then
		ExeSave = false
		exit function
	end if
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.VID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "VENDOR_ID"& Chr(129) & document.all.spanVID.innerText & Chr(129) & "0" & Chr(128)

		sResult = sResult & "NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OWNER_NAME"& Chr(129) & document.all.TxtOwner.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MANAGER_NAME"& Chr(129) & document.all.TxtManager.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_1"& Chr(129) & document.all.TxtAddress1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_2"& Chr(129) & document.all.TxtAddress2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.Txt0City.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.Txt0State.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.Txt0Zip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BILLING_ADDRESS1"& Chr(129) & document.all.TxtBillingAddress1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BILLING_ADDRESS2"& Chr(129) & document.all.TxtBillingAddress2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BILLING_CITY"& Chr(129) & document.all.Txt1City.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BILLING_STATE"& Chr(129) & document.all.Txt1State.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BILLING_ZIP"& Chr(129) & document.all.Txt1Zip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE_DAY"& Chr(129) & document.all.TxtPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE_NIGHT"& Chr(129) & document.all.TxtAfterHoursPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX"& Chr(129) & document.all.TxtFax.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "EMAIL"& Chr(129) & document.all.TxtEmail.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAX_SERVICE_RADIUS"& Chr(129) & document.all.TxtServiceRadius.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE_LICENSE_NUMBER"& Chr(129) & document.all.TxtStateLicense.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY_LICENSE_NUMBER"& Chr(129) & document.all.TxtCityLicense.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE_BACKUP"& Chr(129) & document.all.TxtBackupPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE_BACKUP_TYPE"& Chr(129) & document.all.TxtPhoneType.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PREF_REPORT_METHOD"& Chr(129) & document.all.TxtRepMethod.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOCATION_NUMBER"& Chr(129) & document.all.TxtLocNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FIPS"& Chr(129) & "ZZ7ZZ" & Chr(129) & "1" & Chr(128)
		if document.all.chkEnabled.checked then
			cFlag = "Y"
		else
			cFlag = "N"
		end if
		sResult = sResult & "ENABLED_FLG"& Chr(129) & cFlag & Chr(129) & "1" & Chr(128)
		if document.all.chkPayWithCC.checked then
			cFlag = "Y"
		else
			cFlag = "N"
		end if
		sResult = sResult & "PAY_TYPE_CREDIT"& Chr(129) & cFlag & Chr(129) & "1" & Chr(128)
		if document.all.chkPayWithCheck.checked then
			cFlag = "Y"
		else
			cFlag = "N"
		end if
		sResult = sResult & "PAY_TYPE_CHECK"& Chr(129) & cFlag & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NOTES"& Chr(129) & document.all.TxtNotes.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		if lServiceDaysChangeDetected then
			document.all.VendorServiceDays.Value = getServiceDays
		end if
		if lServiceTypeChangeDetected then
			document.all.ServiceTypes.Value = getServiceTypes
		end if
		document.all.FrmDetails.Submit()
		bRet = true
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If

	ExeSave = bRet
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
	end if
end sub

sub ServiceType_OnChange
	lServiceTypeChangeDetected = true
end sub

sub ServiceDays_OnChange
	lServiceDaysChangeDetected = true
end sub

sub CheckServiceDays(cDay)
	dim cCheckBoxObj, cOpenAtObj, cClosesAtObj

	cCheckBoxObj = "chkOpen" & cDay
	cOpenAtObj = "txtOpensAt" & cDay
	cClosesAtObj = "txtClosesAt" & cDay
	if document.all.item(cCheckBoxObj).checked then
		document.all.item(cOpenAtObj).disabled = false
		document.all.item(cClosesAtObj).disabled = false
	else
		document.all.item(cOpenAtObj).value = ""
		document.all.item(cClosesAtObj).value = ""
		document.all.item(cOpenAtObj).disabled = true
		document.all.item(cClosesAtObj).disabled = true
	end if
	lServiceDaysChangeDetected = true
end sub

sub SetScreenFieldsReadOnly(bReadOnly, strNewClass)

	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnInput") = "TRUE" then
			document.all(iCount).readOnly = bReadOnly
			document.all(iCount).className = strNewClass
		elseif document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next

end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vendor Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="VendorSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" name="ServiceTypes">
<input type="hidden" name="VendorServiceDays">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchVendorID" value="<%=Request.QueryString("SearchVendorID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchAddress" value="<%=Request.QueryString("SearchAddress")%>">
<input type="hidden" name="SearchCity" value="<%=Request.QueryString("SearchCity")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchServiceType" value="<%=Request.QueryString("ServiceType")%>">
<input type="hidden" name="SearchEnabled" value="<%=Request.QueryString("SearchEnabled")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="VID" value="<%=Request.QueryString("VID")%>">
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<%	
dim oConn, oRS, cSQL, VID

VID = Request.QueryString("VID")
If VID <> "" Then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	If VID <> "NEW" then
		cSQL = "SELECT * FROM VENDOR WHERE VENDOR_ID = " & VID
		Set oRS = oConn.Execute(cSQL)
		If Not oRS.EOF then 
			RS_NAME = ReplaceQuotesInText(oRS("NAME"))
			RS_OWNER = oRS("OWNER_NAME")
			RS_MANAGER = ReplaceQuotesInText(oRS("MANAGER_NAME"))
			RS_ADDRESS1 = ReplaceQuotesInText(oRS("ADDRESS_1"))
			RS_ADDRESS2 = ReplaceQuotesInText(oRS("ADDRESS_2"))
			RS_CITY = ReplaceQuotesInText(oRS("CITY"))
			RS_STATE = oRS("STATE")
			RS_ZIP = oRS("ZIP")
			RS_BILLINGADDRESS1 = ReplaceQuotesInText(oRS("BILLING_ADDRESS1"))
			RS__BILLINGADDRESS2 = ReplaceQuotesInText(oRS("BILLING_ADDRESS2"))
			RS_BILLINGCITY = ReplaceQuotesInText(oRS("BILLING_CITY"))
			RS_BILLINGSTATE = oRS("BILLING_STATE")
			RS_BILLINGZIP = oRS("BILLING_ZIP")
			RS_PHONE = oRS("PHONE_DAY")
			RS_BACKUP_PHONE = oRS("PHONE_BACKUP")
			RS_PHONE_TYPE = oRS("PHONE_BACKUP_TYPE")
			RS_PREF_METHOD = oRS("PREF_REPORT_METHOD")
			RS_LOC_NO = oRS("LOCATION_NUMBER")
			RS_AFTERHOURSPHONE = oRS("PHONE_NIGHT")
			RS_FAX = oRS("FAX")
			RS_EMAIL = oRS("EMAIL")
			RS_SERVICERADIUS = oRS("MAX_SERVICE_RADIUS")
			RS_STATELICENSE = oRS("STATE_LICENSE_NUMBER")
			RS_CITYLICENSE = oRS("CITY_LICENSE_NUMBER")
			RS_ENABLED = oRS("ENABLED_FLG")
			RS_PAYWITHCC = oRS("PAY_TYPE_CREDIT")
			RS_PAYWITHCHECK = oRS("PAY_TYPE_CHECK")
			RS_NOTES = ReplaceQuotesInText(oRS("NOTES"))
		end if	
		oRS.Close
		Set oRS = Nothing
	end if	

	if VID <> "NEW" then
		cSQL = "Select * from VENDOR_DAY Where " & _
				"VENDOR_ID=" & VID 
		Set oRS = oConn.Execute(cSQL)
		do while not oRS.eof
			select case oRS("DAY")
				case "MON"
					RS_OPENMON = oRS("OPEN_FLG")
					RS_OPENSATMON = oRS("OPEN_TIME")
					RS_CLOSESATMON = oRS("CLOSE_TIME")
	
				case "TUE"
					RS_OPENTUE = oRS("OPEN_FLG")
					RS_OPENSATTUE = oRS("OPEN_TIME")
					RS_CLOSESATTUE = oRS("CLOSE_TIME")
	
				case "WED"
					RS_OPENWED = oRS("OPEN_FLG")
					RS_OPENSATWED = oRS("OPEN_TIME")
					RS_CLOSESATWED = oRS("CLOSE_TIME")
				
				case "THU"
					RS_OPENTHU = oRS("OPEN_FLG")
					RS_OPENSATTHU = oRS("OPEN_TIME")
					RS_CLOSESATTHU = oRS("CLOSE_TIME")
			
				case "FRI"
					RS_OPENFRI = oRS("OPEN_FLG")
					RS_OPENSATFRI = oRS("OPEN_TIME")
					RS_CLOSESATFRI = oRS("CLOSE_TIME")
			
				case "SAT"
					RS_OPENSAT = oRS("OPEN_FLG")
					RS_OPENSATSAT = oRS("OPEN_TIME")
					RS_CLOSESATSAT = oRS("CLOSE_TIME")
			
				case "SUN"
					RS_OPENSUN = oRS("OPEN_FLG")
					RS_OPENSATSUN = oRS("OPEN_TIME")
					RS_CLOSESATSUN = oRS("CLOSE_TIME")
			
			end select
			oRS.moveNext
		loop
		oRS.close
		set oRS = nothing
	end if
%>		
<table class="LABEL">
<tr>
<td COLSPAN="6">Vendor ID:&nbsp;<span id="spanVID"><%=Request.QueryString("VID")%></span></td>
</tr> 

<tr>
<td >Name:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtName" VALUE="<%=RS_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Owner name:<br><input ScrnInput="TRUE" size="12" CLASS="LABEL" MAXLENGTH="20" TYPE="TEXT" NAME="TxtOwner" VALUE="<%=RS_OWNER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan=4>Manager name:<br><input ScrnInput="TRUE" size="12" CLASS="LABEL" MAXLENGTH="20" TYPE="TEXT" NAME="TxtManager" VALUE="<%=RS_MANAGER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>

<tr>
<td >Address 1:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtAddress1" VALUE="<%=RS_ADDRESS1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Address 2:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtAddress2" VALUE="<%=RS_ADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>ZIP:<br><input size="6" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="Txt0ZIP" VALUE="<%=RS_ZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>City:<br><input CLASS="READONLY" TABINDEX=-1 READONLY MAXLENGTH="20" TYPE="TEXT" NAME="Txt0City" VALUE="<%=RS_CITY%>" ></td>
<td colspan=2>State:<br><input CLASS="READONLY" TABINDEX=-1 READONLY MAXLENGTH="3" SIZE=3 TYPE="TEXT" NAME="Txt0State" VALUE="<%=RS_STATE%>" ></td>
</tr>

<tr> 
<td >Billing address 1:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtBillingAddress1" VALUE="<%=RS_BILLING_ADDRESS1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Billing address 2:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtBillingAddress2" VALUE="<%=RS_BILLING_BADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Bill ZIP:<br><input size="6" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="Txt1ZIP" VALUE="<%=RS_BILLINGZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Billing city:<br><input CLASS="READONLY" TABINDEX=-1 READONLY MAXLENGTH="20" TYPE="TEXT" NAME="Txt1City" VALUE="<%=RS_BILLINGCITY%>" ></td>
<td colspan=2>Billing state:<br><input CLASS="READONLY" TABINDEX=-1 READONLY MAXLENGTH="3" SIZE=3 TYPE="TEXT" NAME="Txt1State" VALUE="<%=RS_BILLINGSTATE%>" ></td>
</tr>

<tr>
<td>Day time Phone:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" TYPE="TEXT" NAME="TxtPhone" VALUE="<%=RS_PHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan=2>Night time phone:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" TYPE="TEXT" NAME="TxtAfterHoursPhone" VALUE="<%=RS_AFTERHOURSPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Fax:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" TYPE="TEXT" NAME="TxtFax" VALUE="<%=RS_FAX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan=2>Email:<br><input size=15 ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="20" TYPE="TEXT" NAME="TxtEmail" VALUE="<%=RS_EMAIL%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>

<tr>
<td>Backup phone:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" TYPE="TEXT" NAME="TxtBackupPhone" VALUE="<%=RS_BEEPERPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan=2>Backup phone type:<br>
	<select STYLE="WIDTH:75" NAME="TxtPhoneType" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		<option VALUE="C">Cellular
		<option VALUE="P">Pager
		<option VALUE="R">Regular
	</select>
</td>
<td colspan=3>Max. service radius:<br><input size=5 ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtServiceRadius" VALUE="<%=RS_SERVICERADIUS%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>

<tr> 
<td >State license no.:<br><input size=10 ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="15" TYPE="TEXT" NAME="TxtStateLicense" VALUE="<%=RS_STATELICENSE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td >City license no.:<br><input size="10" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="15" TYPE="TEXT" NAME="TxtCityLicense" VALUE="<%=RS_CITYLICENSE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td colspan=2>Preferred Rep. method:<br>
	<select STYLE="WIDTH:75" NAME="TxtRepMethod" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		<option VALUE="F">Fax
		<option VALUE="E">Email
		<option VALUE="N">Neither
	</select>
</td>
<td >Location No.:<br><input size="10" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="15" TYPE="TEXT" NAME="TxtLocNumber" VALUE="<%=RS_CITYLICENSE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>

<tr>
<td colspan=3>
<div align="LEFT" style="display:block;height:94px;width:330px;overflow:auto">
	<table border="1" class="LABEL" ID="tblFields" name="tblFields" width="100%" >
	<tr align="left">
		<th ><nobr>Serv. type</th>
		<th ><nobr>Description</th>
		<th ><nobr>Provided</th>
	</tr>
	<tbody ID="TableRows">
	<%
	cSQL = "Select * from SERVICE_TYPE"
	Set oRS = oConn.Execute(cSQL)
	do while not oRS.eof
		nServiceID = oRS("SERVICE_TYPE_ID")
		%>
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" STID='<%=nServiceID%>'>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("TYPE"))%></td>
			<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("DESCRIPTION"))%></td>
			<%
			if VID = "NEW" then
				cChecked = ""
			else
				cSQL = "Select * from VENDOR_SERVICE Where " & _
					"VENDOR_ID=" & VID & " AND SERVICE_TYPE_ID = " & nServiceID
				Set oRS0 = oConn.Execute(cSQL)
				if not oRS0.eof then
					cChecked = "CHECKED"
				else
					cChecked = ""
				end if
				oRS0.close
			end if
			%>
			<td align="center" CLASS="ResultCell"><input type="checkbox" <%=cChecked%> ONCLICK="VBScript::ServiceType_OnChange"></td>
		</tr>
		<%
		oRS.moveNext
	loop
	oRS.close
	set oRS = nothing
	set oRS0 = nothing
	oConn.Close 
	set oConn = nothing
	%>
	</tbody>
	</table>
</div>
</td>

<td colspan="3">
<table border="1" class="LABEL" COLSPAN="4" ID="ServiceDays" name="ServiceDays">
  <tr>
    <td>&nbsp;</td>
    <th align="center">Mon</th>
    <th align="center">Tue</th>
    <th align="center">Wed</th>
    <th align="center">Thu</th>
    <th align="center">Fri</th>
    <th align="center">Sat</th>
    <th align="center">Sun</th>
  </tr>
  <tr>
    <td>Open</td>
    <td align="center"><input type="checkbox" name="chkOpenMon" ONCLICK="VBScript::CheckServiceDays('MON')"></td>
    <td align="center"><input type="checkbox" name="chkOpenTue" ONCLICK="VBScript::CheckServiceDays('TUE')"></td>
    <td align="center"><input type="checkbox" name="chkOpenWed" ONCLICK="VBScript::CheckServiceDays('WED')"></td>
    <td align="center"><input type="checkbox" name="chkOpenThu" ONCLICK="VBScript::CheckServiceDays('THU')"></td>
    <td align="center"><input type="checkbox" name="chkOpenFri" ONCLICK="VBScript::CheckServiceDays('FRI')"></td>
    <td align="center"><input type="checkbox" name="chkOpenSat" ONCLICK="VBScript::CheckServiceDays('SAT')"></td>
    <td align="center"><input type="checkbox" name="chkOpenSun" ONCLICK="VBScript::CheckServiceDays('SUN')"></td>
  </tr>
  <tr>
    <td>Opens at</td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtMon" value="<%=RS_OPENSATMON%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtTue" value="<%=RS_OPENSATTUE%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtWed" value="<%=RS_OPENSATWED%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtThu" value="<%=RS_OPENSATTHU%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtFri" value="<%=RS_OPENSATFRI%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtSat" value="<%=RS_OPENSATSAT%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtOpensAtSun" value="<%=RS_OPENSATSUN%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
  </tr>
  <tr>
    <td>Closes at</td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtMon" value="<%=RS_CLOSESATMON%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtTue" value="<%=RS_CLOSESATTUE%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtWed" value="<%=RS_CLOSESATWED%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtThu" value="<%=RS_CLOSESATTHU%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtFri" value="<%=RS_CLOSESATFRI%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtSat" value="<%=RS_CLOSESATSAT%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
    <td align="center"><input type="text" CLASS="LABEL" name="txtClosesAtSun" value="<%=RS_CLOSESATSUN%>" size="5" ONCHANGE="VBScript::ServiceDays_OnChange"></td>
  </tr>
</table>
</td>
</tr>


<tr>
<td colspan=6>
<input type="checkbox" name="chkPayWithCC">Pay w/Credit card
</td>
</tr>

<tr>
<td colspan=4>
<input type="checkbox" name="chkPayWithCheck">Pay w/Check
</td>
<td colspan="2">
<input type="checkbox" name="chkEnabled">Enabled
</td>
</tr>

<tr>
<td colspan=6>Notes:<br><input size="60" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtNotes" VALUE="<%=RS_NOTES%>"></td>
</tr>
</table>
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Vendor selected.
</div>
<% End If %>
</form>
<script language=vbscript>
Sub window_onload
<%
if len(Request.QueryString("VID")) <> 0 then
	if Request.QueryString("VID")="NEW" then%>
		document.all("chkEnabled").checked = true
	<%else
%>
		document.all("chkEnabled").checked = <%=RS_ENABLED = "Y"%>
		<%
		if not isNull(RS_PAYWITHCHECK) then
		%>
			document.all("chkPayWithCheck").checked = <%=RS_PAYWITHCHECK = "Y"%>
		<%
		end if
		if not isNull(RS_PAYWITHCC) then
		%>
			document.all("chkPayWithCC").checked = <%=RS_PAYWITHCC = "Y"%>
		<%
		end if
		%>
		document.all("TxtBackupPhone").value = "<%=RS_BACKUP_PHONE%>"
		document.all("TxtPhoneType").value = "<%=RS_PHONE_TYPE%>"
		document.all("TxtRepMethod").value = "<%=RS_PREF_METHOD%>"
		document.all("TxtLocNumber").value = "<%=RS_LOC_NO%>"
<%
	end if
	if RS_OPENMON = "Y" then
%>	
		document.all("chkOpenMon").checked = true
<%
	else
%>
		document.all("txtOpensAtMon").disabled = true
		document.all("txtClosesAtMon").disabled = true
<%
	end if
	if RS_OPENTUE = "Y" then
%>	
		document.all("chkOpenTue").checked = true
<%
	else
%>
		document.all("txtOpensAtTue").disabled = true
		document.all("txtClosesAtTue").disabled = true
<%
	end if
	if RS_OPENWED = "Y" then
%>	
		document.all("chkOpenWed").checked = true
<%
	else
%>
		document.all("txtOpensAtWed").disabled = true
		document.all("txtClosesAtWed").disabled = true
<%
	end if
	if RS_OPENTHU = "Y" then
%>	
		document.all("chkOpenThu").checked = true
<%
	else
%>
		document.all("txtOpensAtThu").disabled = true
		document.all("txtClosesAtThu").disabled = true
<%
	end if
	if RS_OPENFRI = "Y" then
%>	
		document.all("chkOpenFri").checked = true
<%
	else
%>
		document.all("txtOpensAtFri").disabled = true
		document.all("txtClosesAtFri").disabled = true
<%
	end if
	if RS_OPENSAT = "Y" then
%>	
		document.all("chkOpenSat").checked = true
<%
	else
%>
		document.all("txtOpensAtSat").disabled = true
		document.all("txtClosesAtSat").disabled = true
<%
	end if
	if RS_OPENSUN = "Y" then
%>	
		document.all("chkOpenSun").checked = true
<%
	else
%>
		document.all("txtOpensAtSun").disabled = true
		document.all("txtClosesAtSun").disabled = true
<%
	end if
	
		if CStr(Request.QueryString("MODE")) = "RO" then %>
			SetScreenFieldsReadOnly true,"DISABLED"
	
	<%	end if %>
<%		if CStr(Request.QueryString("BranchTypeFilter")) <> "" then %>
			SetBranchTypeFieldReadOnly true 
	<%	end if
end if
%>
End Sub
</script>
</body>
</html>


