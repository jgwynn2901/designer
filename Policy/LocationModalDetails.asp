<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%Response.Expires=0%>
<!--#include file="..\lib\ZIP.inc"-->	

<%
dim oConn, oRS, cSQL, cPOL_ID, cAHSID

cPOL_ID = Request.QueryString("PID")
cAHSID = Request.QueryString("AHS_ID")
RSPOLICY_ID = cPOL_ID	
If cAHSID <> "NEW" THEN
	IF cAHSID <> "" then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		cSQL = "Select * " _
			& "From ACCOUNT_HIERARCHY_STEP " _
			& "Where ACCNT_HRCY_STEP_ID = " & cAHSID
		Set oRS = oConn.Execute(cSQL)
		If Not oRS.EOF then
            RSAHS_ID = cAHSID
			RSADDRESS1 = ReplaceQuotesInText(oRS("ADDRESS_1"))
			RSCITY = ReplaceQuotesInText(oRS("CITY"))
			RSSTATE = ReplaceQuotesInText(oRS("STATE"))
			RSZIP = oRS("ZIP")
		end if	
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	End If
End If
%>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Location Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CLocationSearchObj()
{
     this.AHS_ID = "";
	 this.Selected = false;
}

function CPolicySearchObj()
{
	this.PID = "";
	this.Selected = false;
}
var PolicySearchObj = new CPolicySearchObj();
var LocationSearchObj = new CLocationSearchObj();

</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable =  false

Sub window_onload
<%	if Request.QueryString("MODE") = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if cAHS_ID <> "" then %>
			document.all.CoverageFrame.style.height = .3 *  document.body.clientHeight
			document.all.VendorFrame.style.height = .3 *  document.body.clientHeight
<%		end if	
	end if 
%>
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub UpdateAHSID(cAHSID)
	document.all.spanAHS_ID.innerHTML = cAHSID
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Function getAHS_ID
	if document.all.AHS_ID.value <> "NEW" then
		getAHS_ID = document.all.AHS_ID.value
	else
		getAHS_ID = ""
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
	If document.all.TxtStreet.value = "" then
		MsgBox "Street is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If document.all.Zip.value = "" then
		MsgBox "Zip is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	ValidateScreenData = true
End Function


Function ExeSave
	bRet = false
	
	if not ValidateScreenData then 
		ExeSave = false
		exit function
	end if
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.AHS_ID.value = "" then
		ExeSave = false
		exit function
	end if
	
	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.AHS_ID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		'sResult = "ACCNT_HRCY_STEP_ID"& Chr(129) & <%=cAHSID%> & Chr(129) & "1" & Chr(128)	
		'sResult = sResult & "POLICY_ID"& Chr(129) & <%=cPOL_ID%> & Chr(129) & "1" & Chr(128)

		'sResult = sResult & "AHS_POLICY_ID"& Chr(129) & document.all.AHS_POL_ID.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "ADDRESS1"& Chr(129) & document.all.TxtStreet.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "CITY"& Chr(129) & document.all.City.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "ZIP"& Chr(129) & document.all.Zip.value & Chr(129) & "1" & Chr(128)
		'sResult = sResult & "STATE"& Chr(129) & document.all.STATE.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		FrmDetails.Submit()
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
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
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
	
	strURL = "..\Policy\PolicyMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_LOCATION&SELECTONLY=TRUE&PID=" & PID
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


</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Location Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="LocationSave.asp" target="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" name="SearchLID" value="<%=Request.QueryString("SearchLID")%>">
<input type="hidden" name="SearchPOLICY_ID" value="<%=Request.QueryString("SearchPOLICY_ID")%>">

<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="AHS_ID" value="<%=Request.QueryString("AHS_ID")%>">
<input type="hidden" NAME="POLICY_ID" value="<%=Request.QueryString("PID")%>">
<input type="hidden" NAME="LOB" value="<%=Request.QueryString("LOB")%>">

<%If cAHSID <> "" Then %>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td WIDTH="14"><img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count"></td><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">:<span id="SpanSharedCount"><%=SharedCount%></span></td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<span CLASS="LABEL">AHS ID:&nbsp;<span id="spanAHS_ID"><%=cAHSID%></span>

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL" BORDER="0">
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Policy ID:<br><input ScrnInput="TRUE" STYLE="BACKGROUND-COLOR:SILVER" READONLY CLASS="READONLY" MAXLENGTH="10" SIZE="10" TYPE="TEXT" NAME="TxtPOLICY_ID" VALUE="<%=RSPOLICY_ID %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
    <td CLASS="LABEL" COLSPAN="2">Street:<br><input ScrnInput="TRUE" size="40" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtStreet" VALUE="<%=RSADDRESS1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Zip:<br><input CLASS="LABEL" ScrnInput="TRUE" MAXLENGTH="10" size="8" TYPE="TEXT" NAME="Zip" VALUE="<%=RSZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">City:<br><input READONLY CLASS="READONLY" MAXLENGTH="40" size="10" TYPE="TEXT" NAME="City" VALUE="<%=RSCITY%>"></td>
	<td CLASS="LABEL">State:<br><input READONLY CLASS="READONLY" MAXLENGTH="4" size="4" TYPE="TEXT" NAME="State" VALUE="<%=RSSTATE%>"></td>	
	</tr>
	<tr>&nbsp</tr>
</table>

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
<iframe FRAMEBORDER="0" name="CoverageFrame" width="100%" height="100%" src="LocationCoverageClient.asp?<%=Request.QueryString%>"></iframe>
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
<iframe FRAMEBORDER="0" name="VendorFrame" width="100%" height="100%" src="LocationVendorCode.asp?<%=Request.QueryString%>"></iframe>
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Location selected.
</div>
<%End If %>
</form>
</body>
</html>


