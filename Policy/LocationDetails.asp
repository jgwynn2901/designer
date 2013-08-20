<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->

<%	
Response.Expires=0 
Response.Buffer = true
dim cPID, oConn, oRS, cSQL
cPID = Request.QueryString("PID")
RS_LOB = Request.QueryString("LOB")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Location Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CLocationSearchObj()
{
	this.LID = "";
    this.Saved = false;
	this.Selected = false;
}
var LocationSearchObj  = new CLocationSearchObj();
var g_StatusInfoAvailable = false;

</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 100;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 100;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmLocation.action = strURL
	FrmLocation.method = "GET"
	FrmLocation.target = "_parent"	
	FrmLocation.submit
End Sub

Sub UpdatePID(inPID)
	document.all.PID.value = inPID
	document.all.spanPID.innerText = inPID
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

Function GetPID
	GetPID = document.all.PID.value
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

Function ExeSave
	sResult = ""
	bRet = false
	if not ValidateScreenData then 
		ExeSave = false
		exit function
	end if

	' set default form handler
	FrmLocation.action = "LocationSave.asp"
	
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
		If document.all.PID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		'sResult = sResult & "VEHICLE_ID"& Chr(129) & document.all.VID.value & Chr(129) & "0" & Chr(128)
		'sResult = sResult & "POLICY_ID"& Chr(129) & document.all.TxtPOLICY_ID.value & Chr(129) & "0" & Chr(128)
		'sResult = sResult & "VIN"& Chr(129) & document.all.TxtVIN.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "YEAR"& Chr(129) & document.all.TxtYEAR.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "MAKE"& Chr(129) & document.all.TxtMAKE.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "MODEL"& Chr(129) & document.all.TxtMODEL.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "LICENSE_PLATE"& Chr(129) & document.all.TxtLICENSE_PLATE.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "LICENSE_PLATE_STATE"& Chr(129) & document.all.TxtLICENSE_PLATE_STATE.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "REGISTRATION_STATE"& Chr(129) & document.all.TxtREGISTRATION_STATE.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "COLOR"& Chr(129) & document.all.TxtCOLOR.value & Chr(129) & "1" & Chr(128)
		'document.all.TxtSaveData.Value = sResult
		FrmLocation.method = "POST"
		FrmLocation.Submit()
		bRet = true
	Else
		SpanStatus.innerHTML = "Nothing to Save"
	End If
	
	ExeSave = bRet
	
End Function

Function getSelectedAHS_ID
	getSelectedAHS_ID = document.frames("DataFrame").getSelectedAHS_ID
End Function

Function getSelectedAHS_POL_ID
	getSelectedAHS_POL_ID = document.frames("DataFrame").getSelectedAHS_POL_ID
End Function

Sub ExeButtonsNew()
dim cMODE, cPID, cURL
if Not InEditMode Then
	Exit Sub
End If
If document.all.PID.value = "" Or document.all.PID.value = "NEW" Then
	Exit Sub
End If
LocationSearchObj.Selected = false
cMODE = document.body.getAttribute("ScreenMode")
cPID = document.all.PID.value
cURL = "LocationModalMaintenance.asp?SECURITYPRIV=FNSD_POLICY&AHS_ID=NEW&DETAILONLY=TRUE&LOB=<%=RS_LOB%>&PID=<%=Request.querystring("PID")%>"
showModalDialog cURL, LocationSearchObj, "dialogWidth=500px; dialogHeight=570px; center=yes"
If LocationSearchObj.Selected Then
	Refresh
end if
End Sub

Sub ExeButtonsEdit()
dim cMODE, cPID, cURL, cAHS_ID
If Not InEditMode Then
	Exit Sub
End If
cPID = document.all.PID.value
cAHS_ID = getSelectedAHS_ID
If cPID = "" Or cPID = "NEW" Then 
	Exit Sub
end if
If cAHS_ID <> "" Then
	LocationSearchObj.Selected = false
	cMODE = document.body.getAttribute("ScreenMode")
	document.all.AHSID.Value = cAHS_ID
	cURL = "LocationModalMaintenance.asp?SECURITYPRIV=FNSD_POLICY&AHS_ID=" & cAHS_ID & "&DETAILONLY=TRUE&LOB=<%=RS_LOB%>&PID=<%=Request.querystring("PID")%>"
	showModalDialog cURL, LocationSearchObj, "dialogWidth=500px; dialogHeight=570px; center=yes"
Else
	MsgBox "Please select a location to edit.", 0, "FNSNet Designer"		
End If

End Sub

Sub Refresh
dim cPID, cLOB
	
cPID = document.all.PID.value
cLOB = document.all.LOB.value
document.all.tags("IFRAME").item("DataFrame").src = "LocationDetailsData.asp?PID=" & cPID & "&LOB=" & cLOB
End Sub

Sub ExeButtonsRemove
dim cAHS_ID, cAHS_POL_ID
If Not InEditMode Then
	Exit Sub
End If
cAHS_ID = getSelectedAHS_ID
cAHS_POL_ID = getSelectedAHS_POL_ID
If cAHS_ID <> "" and cAHS_POL_ID <> "" Then
	document.all.AHSID.Value = cAHS_ID
	document.all.AHS_POLID.Value = cAHS_POL_ID
	document.all.TxtAction.Value = "DELETE"
	FrmLocation.action = "LocationSave.asp"
	FrmLocation.method = "POST"
	FrmLocation.target = "hiddenPage"	
	FrmLocation.submit
Else
	MsgBox "Please select a location to remove.", 0, "FNSNet Designer"		
End If
End Sub

Function InEditMode
InEditMode = true
	
If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
	MsgBox "This screen is read only.",0,"FNSNetDesigner"
	InEditMode = false
End If

End Function

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
	End If		
End Sub


</script>
<SCRIPT LANGUAGE="JavaScript" FOR="VehicleBtnControl" EVENT="onscriptletevent (event, obj)">
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
		case "REFRESHBUTTONCLICK":
				Refresh();
			break;
		default:
			break;
	}
</SCRIPT>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Location</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmLocation">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchPID" value="<%=Request.QueryString("SearchPID")%>">
<input type="hidden" name="SearchNumber" value="<%=Request.QueryString("SearchNumber")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchCarrier" value="<%=Request.QueryString("SearchCarrier")%>">
<input type="hidden" name="SearchAgent" value="<%=Request.QueryString("SearchAgent")%>">
<input type="hidden" name="SearchLOBCD" value="<%=Request.QueryString("SearchLOBCD")%>">
<input type="hidden" name="SearchMCTYPE" value="<%=Request.QueryString("SearchMCTYPE")%>">
<input type="hidden" name="SearchSelfInsuredFlg" value="<%=Request.QueryString("SearchSelfInsuredFlg")%>">
<input type="hidden" name="SearchEffective" value="<%=Request.QueryString("SearchEffective")%>">
<input type="hidden" name="SearchOriginalEffective" value="<%=Request.QueryString("SearchOriginalEffective")%>">
<input type="hidden" name="SearchExpiration" value="<%=Request.QueryString("SearchExpiration")%>">
<input type="hidden" name="SearchCancellation" value="<%=Request.QueryString("SearchCancellation")%>">
<input type="hidden" name="SearchChange" value="<%=Request.QueryString("SearchChange")%>">
<input type="hidden" name="SearchLoad" value="<%=Request.QueryString("SearchLoad")%>">
<input type="hidden" name="SearchCompanyCode" value="<%=Request.QueryString("SearchCompanyCode")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PID" value="<%=Request.QueryString("PID")%>">
<input type="hidden" NAME="LOB" value="<%=Request.QueryString("LOB")%>">
<input type="hidden" NAME="AHSID">
<input type="hidden" NAME="AHS_POLID">

<%If cPID <> "" Then%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" >
<tr>
<tr>
<tr>
<tr>
<tr>
<td>Policy ID:&nbsp;<span id="spanPID"><%=Request.QueryString("PID")%></span></td>

</tr>
<tr>
</table>

<span class="Label" ID=SPANDATA>Retrieving...</span>
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?NEWCAPTION=Add&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=FALSE&HIDEATTACH=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=VehicleBtnControl type=text/x-scriptlet></OBJECT>
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="LocationDetailsData.asp?<%=Request.QueryString%>&AHSID=<%=RSAHSID%>">
</fieldset>
<%else%>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No policy selected.
</div>
<% End If %>
</form>
</body>
</html>


