 <!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%Response.Expires=0 
Dim cVID, cPID, oConn, oRS, cSQL

cVID = Request.QueryString("VID")
cPID = Request.QueryString("PID")
If cVID <> "" Then
	If cVID <> "NEW" then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		cSQL = "SELECT * FROM VEHICLE WHERE VEHICLE_ID = " & cVID
		Set oRS = oConn.Execute(cSQL)
		If Not oRS.EOF then
			RSPOLICY_ID = oRS("POLICY_ID")
			RSVIN = ReplaceQuotesInText(oRS("VIN"))
			RSYEAR = ReplaceQuotesInText(oRS("YEAR"))
			RSMAKE = ReplaceQuotesInText(oRS("MAKE"))
			RSMODEL = ReplaceQuotesInText(oRS("MODEL"))
			RSLICENSE_PLATE = ReplaceQuotesInText(oRS("LICENSE_PLATE"))
			RSLICENSE_PLATE_STATE = oRS("LICENSE_PLATE_STATE")
			RSREGISTRATION_STATE = oRS("REGISTRATION_STATE")
			RSCOLOR = ReplaceQuotesInText(oRS("COLOR"))
		end if	
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	end if	
End If
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vehicle Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CDriverSearchObj()
{
	this.DID = "";
	this.Selected = false;
	this.Saved = false;
}
function CCoverageSearchObj()
{
	//this.COVID = "";
	this.Selected = false;	
	this.Saved = false;	
}

function CPolicySearchObj()
{
	this.PID = "";
	this.Selected = "";
}
var PolicySearchObj = new CPolicySearchObj();
var DriverSearchObj = new CDriverSearchObj();

</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable =  false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if cVID <> "" then %>
		document.all.TxtREGISTRATION_STATE.Value = "<%= RSREGISTRATION_STATE %>"
		document.all.TxtLICENSE_PLATE_STATE.Value = "<%= RSLICENSE_PLATE_STATE %>"
<%		end if	
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "VehicleSearch-f.asp"
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

Function GetVIDName
	GetVIDName = document.all.TxtName.value
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
dim cMsg

cMsg = ""
ValidateScreenData = true

If document.all.TxtVIN.value = "" then
	cMsg = "VIN is a required field." & vbCRLF
end if
If document.all.TxtYEAR.value = "" then
	cMsg = cMsg & "Year is a required field." & vbCRLF
end if
If document.all.TxtMAKE.value = "" then
	cMsg = cMsg & "Make is a required field." & vbCRLF
end if
If document.all.TxtMODEL.value = "" then
	cMsg = cMsg & "Model is a required field." & vbCRLF
end if
if len( cMsg ) <> 0 then
	MsgBox cMsg, 0, "FNSNetDesigner"
	ValidateScreenData = false
end if
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.", 0, "FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.VID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.VID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false

	if not ValidateScreenData then 
		ExeSave = false
		exit function
	end if

	' set default form handler
	FrmDetails.action = "VehicleSave.asp"
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.", 0, "FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.VID.value = "" then
		ExeSave = false
		exit function
	end if
	
	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.VID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		sResult = sResult & "VEHICLE_ID"& Chr(129) & document.all.VID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "POLICY_ID"& Chr(129) & document.all.TxtPOLICY_ID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "VIN"& Chr(129) & document.all.TxtVIN.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "YEAR"& Chr(129) & document.all.TxtYEAR.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAKE"& Chr(129) & document.all.TxtMAKE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MODEL"& Chr(129) & document.all.TxtMODEL.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LICENSE_PLATE"& Chr(129) & document.all.TxtLICENSE_PLATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LICENSE_PLATE_STATE"& Chr(129) & document.all.TxtLICENSE_PLATE_STATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "REGISTRATION_STATE"& Chr(129) & document.all.TxtREGISTRATION_STATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "COLOR"& Chr(129) & document.all.TxtCOLOR.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	Else
		SpanStatus.innerHTML = "Nothing to Save"
	End If
	
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

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.", 0, "FNSNetDesigner"	
	End If		
End Sub


Function GetSelectedDID
	GetSelectedDID = document.frames("DriverFrame").GetSelectedDID
End Function

Function InEditMode

	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "This screen is read only.", 0, "FNSNetDesigner"
		InEditMode = false
	End If
End Function

Sub ExeButtonsNew
dim cVID
dim cMODE, cURL
If Not InEditMode Then
	Exit Sub
End If
cVID = document.all.VID.value
If cVID = "" Or cVID = "NEW" Then	
	Exit Sub
end if
DriverSearchObj.DID = ""
cMODE = document.body.getAttribute("ScreenMode")
cURL = "DriverMaintenance.asp?SECURITYPRIV=FNSD_VEHICLE&PID=<%=cPID%>&DID=NEW&CONTAINERTYPE=MODAL"
showModalDialog cURL, DriverSearchObj, "center"
if len(DriverSearchObj.DID) <> 0 Then
	multi = Replace(DriverSearchObj.DID,"||",",")
	'self.location.href = "AHBranchSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & multi
End If

If DriverSearchObj.Saved Then 
	Refresh
end if
End Sub

Sub ExeButtonsEdit
dim cVID, cDID, cMODE, cURL
	
If Not InEditMode Then
	Exit Sub
End If
cVID = document.all.VID.value
If cVID = "" Or cVID = "NEW" Then 
	Exit Sub
end if
cDID = GetSelectedDID
If cDID <> "" Then
	DriverSearchObj.Saved = false
	cMODE = document.body.getAttribute("ScreenMode")
	cURL = "DriverMaintenance.asp?SECURITYPRIV=FNSD_VEHICLE&VID=" & cVID & "&DID=" & cDID & "&CONTAINERTYPE=MODAL"
	showModalDialog  cURL, DriverSearchObj, "center"
	If DriverSearchObj.Saved Then 
		Refresh
	end if
Else
	MsgBox "Please select a driver to edit.", 0, "FNSNet Designer"		
End If
End Sub

Sub Refresh
	document.all.tags("IFRAME").item("DriverFrame").src = "VehicleDriverDetails.asp?PID=<%=cPID%>"
End Sub

Sub ExeButtonsRemove
	If Not InEditMode Then
		Exit Sub
	End If
   dim DID, sResult
	DID = GetSelectedDID
	If DID <> "" Then
	    sResult = sResult &  DID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmDetails.action = "DriverSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a driver to remove.", 0, "FNSNet Designer"		
	End If

End Sub

Function AttachPolicy (ID)
	PID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	PolicySearchObj.PID = PID
	PolicySearchObj.Selected = false

	If PID = "" Then PID = "NEW"
	
	If PID = "NEW" And MODE = "RO" Then
		MsgBox "No policy currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Policy\PolicyMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_VEHICLE&SELECTONLY=TRUE&PID=" & PID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,PolicySearchObj ,"center"

	'if Selected=true update everything
	If PolicySearchObj.Selected = true Then
		If PolicySearchObj.PID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = PolicySearchObj.PID
		end if
	End If
End Function

<!--#include file="..\lib\Help.asp"-->
</script>
<script LANGUAGE="JavaScript" FOR="DriverBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "REMOVEBUTTONCLICK":
				ExeButtonsRemove();
			break;
		case "EDITBUTTONCLICK":
				ExeButtonsEdit();
			break;
		case "NEWBUTTONCLICK":
				ExeButtonsNew();
			break;
		default:
			break;
	}
</script>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="14" HEIGHT="10"><nobr>&nbsp;» Vehicle Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
</td></tr>
<tr><td CLASS="GrpLabelLine" colspan="1" HEIGHT="1"></td></tr>
<tr><td colspan="1" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="VehicleSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" name="SearchVID" value="<%=Request.QueryString("SearchVID")%>">
<input type="hidden" name="SearchPOLICY_ID" value="<%=Request.QueryString("SearchPOLICY_ID")%>">
<input type="hidden" name="SearchVIN" value="<%=Request.QueryString("SearchVIN")%>">
<input type="hidden" name="SearchYEAR" value="<%=Request.QueryString("SearchYEAR")%>">
<input type="hidden" name="SearchMAKE" value="<%=Request.QueryString("SearchMAKE")%>">
<input type="hidden" name="SearchMODEL" value="<%=Request.QueryString("SearchMODEL")%>">
<input type="hidden" name="SearchLICENSE_PLATE" value="<%=Request.QueryString("SearchLICENSE_PLATE")%>">
<input type="hidden" name="SearchLICENSE_PLATE_STATE" value="<%=Request.QueryString("SearchLICENSE_PLATE_STATE")%>">
<input type="hidden" name="SearchREGISTRATION_STATE" value="<%=Request.QueryString("SearchREGISTRATION_STATE")%>">
<input type="hidden" name="SearchCOLOR" value="<%=Request.QueryString("SearchCOLOR")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="VID" value="<%=Request.QueryString("VID")%>">

<%If cVID <> "" Then %>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<!--<td WIDTH="14"><img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count"></td><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">:<span id="SpanSharedCount"><%=SharedCount%></span></td>-->
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<span CLASS="LABEL">Vehicle ID:&nbsp;<span id="spanVID"><%=Request.QueryString("VID")%></span>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL" BORDER="0">
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Policy ID:<br><input ScrnInput="TRUE" READONLY STYLE="BACKGROUND-COLOR:SILVER" CLASS="LABEL" MAXLENGTH="10" SIZE="10" TYPE="TEXT" NAME="TxtPOLICY_ID" VALUE="<%=cPID %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL" COLSPAN="2">VIN:<br><input ScrnInput="TRUE" size="40" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtVIN" VALUE="<%=RSVIN%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Year:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="4" size="4" TYPE="TEXT" NAME="TxtYEAR" VALUE="<%=RSYEAR%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL"><nobr>Registration State:<br>
	<select ScrnBtn="TRUE" NAME="TxtREGISTRATION_STATE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	<option VALUE>
	<!--#include file="..\lib\States.asp"-->
	</select>
	</td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Make:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" size="40" TYPE="TEXT" NAME="TxtMAKE" VALUE="<%=RSMAKE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL" COLSPAN="2">Model:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" size="40" TYPE="TEXT" NAME="TxtMODEL" VALUE="<%=RSMODEL%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">License Plate:<br><input ScrnInput="TRUE" size="30" TYPE="TEXT" MAXLENGTH="30" NAME="TxtLICENSE_PLATE" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSLICENSE_PLATE%>"></td>
	<td CLASS="LABEL">Color:<br><input ScrnInput="TRUE" size="30" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtCOLOR" VALUE="<%=RSCOLOR%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL"><nobr>License Plate State:<br>
	<select ScrnBtn="TRUE" NAME="TxtLICENSE_PLATE_STATE" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	<option VALUE>
	<!--#include file="..\lib\States.asp"-->
	</select>
	</td>
	</tr>
</table>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Drivers</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'70';width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="DriverBtnControl" type="text/x-scriptlet"></object>
<iframe FRAMEBORDER="0" height="70" width="100%" name="DriverFrame" src="VehicleDriverDetails.asp?<%=Request.QueryString%>"></iframe>
</fieldset>
<br>

<table CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Coverage Codes Received From Client</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="175" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset ID="fldSet2" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<iframe FRAMEBORDER="0" name="CoverageFrame" SCROLLING=AUTO width="100%" height="76%" src="VehicleCoverageDetails.asp?<%=Request.QueryString%>"></iframe>
</fieldset>

<br>
<table CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vendor Designators Created by XREF</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="175" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset ID="fldSet3" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<iframe FRAMEBORDER="0" name="VendorFrame" SCROLLING=AUTO width="100%" height="76%" src="VendorDataDetails.asp?<%=Request.QueryString%>"></iframe>
</fieldset>
<%Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No vehicle selected.
</div>
<%End If%>
</form>
</body>
</html>


