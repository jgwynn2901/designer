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
sub fillZip1
dim cURL

oZip.Zip = ""
oZip.State = ""
oZip.City = ""
cURL = "../LIB/ZIPLookupIFrame.asp?ZIP=" & document.all.TxtOZip.value
showModalDialog  cURL, oZip ,"center"
document.all.TxtOZip.value = oZip.Zip
document.all.TxtOState.value = oZip.State
document.all.TxtOCity.value = oZip.City
end sub

sub TxtOZip_onblur
dim x, lFocusSet

fillZIP1
on error resume next
for x=0 to document.all.length-1
	if uCase(Trim(document.all(x).name)) = "TXTOZIP" then
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

Sub window_onload
dim cInnerHTML

<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
	
<%	end if %>
<%	if CStr(Request.QueryString("BranchTypeFilter")) <> "" then %>
	SetBranchTypeFieldReadOnly true 
<%	end if %>
<%
if Session("lIsCISG") then
%>
	cInnerHTML = document.all.OfficeNo.innerHTML
	cInnerHTML = replace(cInnerHTML, "Office Number", "Manager Code")
	document.all.OfficeNo.innerHTML = cInnerHTML
<%	
	if Request.QueryString("BID") = "NEW" then
%>	
		document.all.TxtAHLoadID.value = "15"
		document.all.TxtBranchType.value = "CLAIMHANDLING"
<%
	end if
	if len(Request.QueryString("BID")) <> 0 then
%>
		document.all.TxtAHLoadID.disabled = true
		document.all.TxtBranchType.disabled = true
<%
	end if
end if
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "BranchSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateBID(inBID)
	document.all.BID.value = inBID
	document.all.spanBID.innerText = inBID
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

Function GetBID
	if document.all.BID.value <> "NEW" then
		GetBID = document.all.BID.value
	else
		GetBID = ""
	end if 
End Function

Function GetBIDOfficeName
	GetBIDOfficeName = document.all.TxtOfficeName.value
End Function

Function GetBNUM
	GetBNUM = document.all.TxtBranchNumber.value
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
	If  document.all.TxtAHLoadID.value = "" then
		MsgBox "AH Load ID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.TxtBranchType.value = "" then
		MsgBox "Branch type is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	if IsNumeric(document.all.TxtAHLoadID.value) = false then
		MsgBox "Please enter a number in the AH Load ID field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if

	ValidateScreenData = true
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.BID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.BID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.BID.value = "" then
		ExeSave = false
		exit function
	end if
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.BID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "BRANCH_ID"& Chr(129) & document.all.spanBID.innerText & Chr(129) & "1" & Chr(128)

		sResult = sResult & "BRANCH_NUMBER"& Chr(129) & document.all.TxtBranchNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCOUNT_HIERARCHY_LOAD_ID"& Chr(129) & document.all.TxtAHLoadID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATUS"& Chr(129) & document.all.TxtStatus.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OFFICE_NUMBER"& Chr(129) & document.all.TxtOfficeNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OFFICE_TYPE"& Chr(129) & document.all.TxtOfficeType.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OFFICE_NAME"& Chr(129) & document.all.TxtOfficeName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_1"& Chr(129) & document.all.TxtAddress1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_2"& Chr(129) & document.all.TxtAddress2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.City.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.State.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.Zip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "COUNTRY"& Chr(129) & document.all.TxtCountry.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FOREIGN_ZIP"& Chr(129) & document.all.TxtForeignZip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ALT_PHONE"& Chr(129) & document.all.TxtAltPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX"& Chr(129) & document.all.TxtFax.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BACKUP_FAX"& Chr(129) & document.all.TxtBackupFax.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OVERNIGHT_ADDRESS1"& Chr(129) & document.all.TxtOAddress1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OVERNIGHT_ADDRESS2"& Chr(129) & document.all.TxtOAddress2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OVERNIGHT_ADDRESS_CITY"& Chr(129) & document.all.TxtOCity.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OVERNIGHT_ADDRESS_STATE"& Chr(129) & document.all.TxtOState.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OVERNIGHT_ADDRESS_ZIP"& Chr(129) & document.all.TxtOZip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CONTACT_F_NAME"& Chr(129) & document.all.TxtContactFName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CONTACT_L_NAME"& Chr(129) & document.all.TxtContactLName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ALT_CONTACT_F_NAME"& Chr(129) & document.all.TxtAlternateContactFName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ALT_CONTACT_L_NAME"& Chr(129) & document.all.TxtAlternateContactLName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CONTACT_TITLE"& Chr(129) & document.all.TxtContactTitle.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NOTES"& Chr(129) & document.all.TxtNotes.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LAT"& Chr(129) & document.all.TxtLat.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LON"& Chr(129) & document.all.TxtLon.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BRANCH_TYPE"& Chr(129) & document.all.TxtBranchType(document.all.TxtBranchType.selectedIndex).value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "EMAIL_ADDRESS"& Chr(129) & document.all.TxtEMail.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
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

sub SetBranchTypeFieldReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("SpecialFilterBtn") = "TRUE" then
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Branch Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="BranchSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchBID" value="<%=Request.QueryString("SearchBID")%>">
<input type="hidden" name="SearchBranchNumber" value="<%=Request.QueryString("SearchBranchNumber")%>">
<input type="hidden" name="SearchAHLoadID" value="<%=Request.QueryString("SearchAHLoadID")%>">
<input type="hidden" name="SearchStatus" value="<%=Request.QueryString("SearchStatus")%>">
<input type="hidden" name="SearchOfficeNumber" value="<%=Request.QueryString("SearchOfficeNumber")%>">
<input type="hidden" name="SearchOfficeType" value="<%=Request.QueryString("SearchOfficeType")%>">
<input type="hidden" name="SearchOfficeName" value="<%=Request.QueryString("SearchOfficeName")%>">
<input type="hidden" name="SearchAddress" value="<%=Request.QueryString("SearchAddress")%>">
<input type="hidden" name="SearchCity" value="<%=Request.QueryString("SearchCity")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="BID" value="<%=Request.QueryString("BID")%>">
<input type="hidden" NAME="SearchBranchType" value="<%=Request.QueryString("SearchBranchType")%>">
<input type="hidden" NAME="BranchTypeFilter" value="<%=Request.QueryString("BranchTypeFilter")%>">
<%	

BID = CStr(Request.QueryString("BID"))
If BID <> "" Then
	If BID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM BRANCH WHERE BRANCH_ID = " & BID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then 
			RSBRANCH_NUMBER = ReplaceQuotesInText(RS("BRANCH_NUMBER"))
			RSACCOUNT_HIERARCHY_LOAD_ID = RS("ACCOUNT_HIERARCHY_LOAD_ID")
			RSSTATUS = ReplaceQuotesInText(RS("STATUS"))
			RSOFFICE_NUMBER = RS("OFFICE_NUMBER")
			RSOFFICE_TYPE = ReplaceQuotesInText(RS("OFFICE_TYPE"))
			RSOFFICE_NAME = ReplaceQuotesInText(RS("OFFICE_NAME"))
			RSADDRESS_1 = ReplaceQuotesInText(RS("ADDRESS_1"))
			RSADDRESS_2 = ReplaceQuotesInText(RS("ADDRESS_2"))
			RSCITY = ReplaceQuotesInText(RS("CITY"))
			RSSTATE = ReplaceQuotesInText(RS("STATE"))
			RSZIP = ReplaceQuotesInText(RS("ZIP"))
			RSCOUNTRY = ReplaceQuotesInText(RS("COUNTRY"))
			RSFOREIGN_ZIP = ReplaceQuotesInText(RS("FOREIGN_ZIP"))
			RSPHONE = ReplaceQuotesInText(RS("PHONE"))
			RSALT_PHONE = ReplaceQuotesInText(RS("ALT_PHONE"))
			RSFAX = ReplaceQuotesInText(RS("FAX"))
			RSBACKUP_FAX = ReplaceQuotesInText(RS("BACKUP_FAX"))
			RSOVERNIGHT_ADDRESS1 = ReplaceQuotesInText(RS("OVERNIGHT_ADDRESS1"))
			RSOVERNIGHT_ADDRESS2 = ReplaceQuotesInText(RS("OVERNIGHT_ADDRESS2"))
			RSOVERNIGHT_ADDRESS_CITY = ReplaceQuotesInText(RS("OVERNIGHT_ADDRESS_CITY"))
			RSOVERNIGHT_ADDRESS_STATE = ReplaceQuotesInText(RS("OVERNIGHT_ADDRESS_STATE"))
			RSOVERNIGHT_ADDRESS_ZIP = ReplaceQuotesInText(RS("OVERNIGHT_ADDRESS_ZIP"))
			RSCONTACT_F_NAME = ReplaceQuotesInText(RS("CONTACT_F_NAME"))
			RSCONTACT_L_NAME = ReplaceQuotesInText(RS("CONTACT_L_NAME"))
			RSALT_CONTACT_F_NAME = ReplaceQuotesInText(RS("ALT_CONTACT_F_NAME"))
			RSALT_CONTACT_L_NAME = ReplaceQuotesInText(RS("ALT_CONTACT_L_NAME"))
			RSCONTACT_TITLE = ReplaceQuotesInText(RS("CONTACT_TITLE"))
			RSNOTES = ReplaceQuotesInText(RS("NOTES"))
			RSLAT = RS("LAT")
			RSLON = RS("LON")
			RSBRANCH_TYPE = ReplaceQuotesInText(RS("BRANCH_TYPE"))
			RSEMAIL = ReplaceQuotesInText(RS("EMAIL_ADDRESS"))
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
%>		
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

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table class="LABEL">
	<tr>
	<tr>
	<tr>
	<tr>
	<td COLSPAN="4">Branch ID:&nbsp;<span id="spanBID"><%=Request.QueryString("BID")%></span></td>
	<td> 
	</tr> 
	<tr>
	<td COLSPAN="2">Branch Number:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtBranchNumber" VALUE="<%=RSBRANCH_NUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>AH Load ID:<br><input ScrnInput="TRUE" size="12" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtAHLoadID" VALUE="<%=RSACCOUNT_HIERARCHY_LOAD_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Status:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtStatus" VALUE="<%=RSSTATUS%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Office Type:<br><input ScrnInput="TRUE" size="6" CLASS="LABEL" MAXLENGTH="5" TYPE="TEXT" NAME="TxtOfficeType" VALUE="<%=RSOFFICE_TYPE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td COLSPAN="4">Office Name:<br><input ScrnInput="TRUE" size="80" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtOfficeName" VALUE="<%=RSOFFICE_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td id="OfficeNo">Office Number:<br><input size="12" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtOfficeNumber" VALUE="<%=RSOFFICE_NUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr> 
	<td COLSPAN="2">Address 1:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtAddress1" VALUE="<%=RSADDRESS_1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Address 2:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" NAME="TxtAddress2" VALUE="<%=RSADDRESS_2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td>Zip:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="9" TYPE="TEXT" NAME="Zip" VALUE="<%=RSZIP%>" ></td>
	<td>City:<br><input CLASS="READONLY" TABINDEX=-1 READONLY MAXLENGTH="20" TYPE="TEXT" NAME="City" VALUE="<%=RSCITY%>" ></td>
	<td>State:<br><input CLASS="READONLY" TABINDEX=-1 READONLY MAXLENGTH="3" SIZE=3 TYPE="TEXT" NAME="State" VALUE="<%=RSSTATE%>" ></td>
	<td>Country:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtCountry" VALUE="<%=RSCOUNTRY%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Foreign Zip:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtForeignZip" VALUE="<%=RSFOREIGN_ZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td COLSPAN="2">Phone:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="14" TYPE="TEXT" NAME="TxtPhone" VALUE="<%=RSPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Alt Phone:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="14" TYPE="TEXT" NAME="TxtAltPhone" VALUE="<%=RSALT_PHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Fax:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFax" VALUE="<%=RSFAX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Backup Fax:<br><input size="16" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtBackupFax" VALUE="<%=RSBACKUP_FAX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td COLSPAN="2">Contact F. Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtContactFName" VALUE="<%=RSCONTACT_F_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Contact L. Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtContactLName" VALUE="<%=RSCONTACT_L_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td NOWRAP>Alt Contact F. Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtAlternateContactFName" VALUE="<%=RSALT_CONTACT_F_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td NOWRAP>Alt Contact L. Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtAlternateContactLName" VALUE="<%=RSALT_CONTACT_L_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr> 
	<td COLSPAN="2">Contact Title:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtContactTitle" VALUE="<%=RSCONTACT_TITLE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>LAT:<br><input size="10" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtLat" VALUE="<%=RSLAT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>LON:<br><input size="10" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtLon" VALUE="<%=RSLON%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr> 
	<td COLSPAN="2">Overnight Address 1:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtOAddress1" VALUE="<%=RSOVERNIGHT_ADDRESS1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Overnight Address 2:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtOAddress2" VALUE="<%=RSOVERNIGHT_ADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr> 
	<td>Ovr. Zip:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="9" TYPE="TEXT" NAME="TxtOZip" VALUE="<%=RSOVERNIGHT_ADDRESS_ZIP%>" ></td>
	<td>Overnight City:<br><input CLASS="READONLY" READONLY TABINDEX=-1 MAXLENGTH="40" TYPE="TEXT" NAME="TxtOCity" VALUE="<%=RSOVERNIGHT_ADDRESS_CITY%>" ></td>	
	<td COLSPAN="2">Overnight State:<br><input CLASS="READONLY" READONLY TABINDEX=-1 MAXLENGTH="3" SIZE=3 TYPE="TEXT" NAME="TxtOState" VALUE="<%=RSOVERNIGHT_ADDRESS_STATE%>" ></td>	
	<td>Branch Type:<br><select SpecialFilterBtn="TRUE" ScrnBtn="TRUE" NAME="TxtBranchType" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><option VALUE></option><option VALUE="CLAIMHANDLING">CLAIMHANDLING</option><option VALUE="MANAGEDCARE">MANAGEDCARE</option></select></td>
	</tr>
	<tr>
	<td COLSPAN="3">Email address:<br><input ScrnInput="TRUE" size="60" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtEMail" VALUE="<%=RSEMAIL%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>		
	<td COLSPAN="2">Notes:<br><input ScrnInput="TRUE" size="60" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtNotes" VALUE="<%=RSNOTES%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>	
	</tr>
	</table>
</td>
</tr> 
</table>

<%		
		If Not IsNull(RSBRANCH_TYPE) Then
			If  CStr(RSBRANCH_TYPE) <> "" Then %>
<script LANGUAGE="VBScript">
	SelectOption document.all.TxtBranchType,"<%=CStr(RSBRANCH_TYPE)%>"
</script>
<%			End If
		End If


		If Not IsNull(RSOVERNIGHT_ADDRESS_STATE) Then
			If  CStr(RSOVERNIGHT_ADDRESS_STATE) <> "" Then %>
<script LANGUAGE="VBScript">
	SelectOption document.all.TxtOState,"<%=CStr(RSOVERNIGHT_ADDRESS_STATE)%>"
</script>
<%			End If
		End If  %>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No branch selected.
</div>


<% End If %>

</form>
</body>
</html>


