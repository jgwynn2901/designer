 <!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->

<%	Response.Expires=0 
	Response.Buffer = true

	RSAHSID	= Request.QueryString("AHSID")

	Dim PID
	PID	= CStr(Request.QueryString("PID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Vehicle</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CVehicleSearchObj()
{
	this.Selected = false;
}
var VehicleSearchObj = new CVehicleSearchObj();
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
	FrmVehicle.action = strURL
	FrmVehicle.method = "GET"
	FrmVehicle.target = "_parent"	
	FrmVehicle.submit
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

Function GetSelectedVID

	GetSelectedVID = document.frames("DataFrame").GetSelectedVID
End Function

Sub ExeButtonsNew()

	If Not InEditMode Then
		Exit Sub
	End If
   
	If document.all.PID.value = "" Or document.all.PID.value = "NEW" Then
		Exit Sub
	End If
	VehicleSearchObj.Selected = false
    dim MODE, PID
	MODE = document.body.getAttribute("ScreenMode")
	PID = document.all.PID.value
	strURL = "VehicleMaintenance.asp?SECURITYPRIV=FNSD_POLICY&PID=" & PID & "&VID=NEW&DETAILONLY=TRUE"
	showModalDialog strURL, VehicleSearchObj, "dialogWidth=500px; dialogHeight=610px; center=yes"	
	If VehicleSearchObj.Selected Then 
		Refresh
	end if
End Sub

Sub ExeButtonsEdit()
dim MODE, PID

	If Not InEditMode Then
		Exit Sub
	End If
  PID = document.all.PID.value
  VID = GetSelectedVID
   If PID = "" Or PID = "NEW" Then Exit Sub
       If VID <> "" Then
      
        VehicleSearchObj.Selected = false
		MODE = document.body.getAttribute("ScreenMode")
		strURL = "VehicleMaintenance.asp?SECURITYPRIV=FNSD_POLICY&PID=" & PID & "&VID=" & VID & "&DETAILONLY=TRUE"
		showModalDialog strURL, VehicleSearchObj, "dialogWidth=500px; dialogHeight=610px; center=yes"
		If VehicleSearchObj.Selected Then 
			Refresh
		end if
	Else
		MsgBox "Please select a vehicle to edit.", 0, "FNSNet Designer"		
	End If

End Sub

Sub Refresh
	PID = document.all.PID.value
	document.all.tags("IFRAME").item("DataFrame").src = "PolicyVehicleData.asp?PID=" & PID
End Sub

Sub ExeButtonsRemove
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.PID.value = "" Or document.all.PID.value = "NEW" Then
		Exit Sub
	End If

	dim VID, sResult
	VID = GetSelectedVID
	
	If VID <> "" Then
	
		sResult = sResult &  VID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmVehicle.action = "VehicleSave.asp"
		FrmVehicle.method = "POST"
		FrmVehicle.target = "hiddenPage"	
		FrmVehicle.submit
		Refresh
	Else
		MsgBox "Please select a vehicle to remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub

Function InEditMode
	InEditMode = true
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "This screen is read only.",0,"FNSNetDesigner"
		InEditMode = false
	End If

End Function

Function ExeCopy
	If Not InEditMode Then
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.PID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
		
	FrmVehicle.action = "PolicyCopy.asp"
	FrmVehicle.method = "GET"
	FrmVehicle.target = "hiddenPage"	
	FrmVehicle.submit
'	Refresh is done inside UserCopy.asp due to timing
	ExeCopy = true
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
		default:
			break;
	}
</SCRIPT>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Policy Vehicle</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmVehicle">
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

<%	
If PID <> "" Then
	If PID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		' DMS: 2/17/00 Changed the SQL to grab the column from 
		'              AHS_POLICY as the column has been removed from 
		'              the POLICY table.

		SQLST = "SELECT AHS_POLICY.ACCNT_HRCY_STEP_ID " &_
				"  FROM AHS_POLICY " &_
				" WHERE Policy_ID = " & PID 
		
		Set RS = Conn.Execute(SQLST)
		
		If Not RS.EOF Then
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
		End If

		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing

	End If

%>
<input type="HIDDEN" NAME="TxtAHSID" value="<%=RSAHSID%>">

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
<OBJECT data="../Scriptlets/ObjButtons.asp?NEWCAPTION=Add&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDEATTACH=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=VehicleBtnControl type=text/x-scriptlet></OBJECT>
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="PolicyVehicleData.asp?<%=Request.QueryString%>">
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No policy selected.
</div>


<% End If %>

</form>
</body>
</html>


