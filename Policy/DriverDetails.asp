<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ZIP.inc"-->
<!--#include file="..\lib\CheckSharedAttribute.inc"-->
<%
Dim SharedCount, SharedCountText, DID
SharedCount = 0
SharedCountText = "Ready"
	
DID	= CStr(Request.QueryString("DID"))
If DID <> "" Then
	If DID = "NEW" Then 
		SharedCount = 0
	Else
		SharedCount = CheckSharedAttribute(CLng(DID),True,True,1,False,False,0)
	End If
End If	
	
RSPOLICY_ID = Request.QueryString("PID")
If DID <> "" Then
	If DID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM DRIVER WHERE DRIVER_ID = " & DID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSSSN = RS("SSN")
			RSPOLICY_ID = RS("POLICY_ID")
			RSNAME_FIRST = ReplaceQuotesInText(RS("NAME_FIRST"))
			RSNAME_LAST = ReplaceQuotesInText(RS("NAME_LAST"))
			RSADDRESS1 = ReplaceQuotesInText(RS("ADDRESS1"))
			RSADDRESS2 = ReplaceQuotesInText(RS("ADDRESS2"))
			RSCITY = RS("CITY")
			RSSTATE = RS("STATE")
			RSZIP = RS("ZIP")
			RSPHONE = RS("PHONE")
			RSRELATION_TO_INSURED = ReplaceQuotesInText(RS("RELATION_TO_INSURED"))
			RSLICENSE_NUMBER = RS("LICENSE_NUMBER")
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
End If
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Carrier Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if DID <> "" then %>
		document.all.STATE.Value = "<%= RSSTATE %>"
			<% if SharedCount <= 1 then %>
<%	else %>
	SetStatusInfoAvailableFlag(true)
<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "DriverSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateDID(inDID)
	document.all.DID.value = inDID
	document.all.spanDID.innerText = inDID
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

Function GetDID
	if document.all.DID.value <> "NEW" then
		GetDID = document.all.DID.value
	else
		GetDID = ""
	end if 
End Function

Function GetDIDName
	GetDIDName = document.all.TxtName.value
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

ValidateScreenData = true

If document.all.TxtNAME_FIRST.value = "" then
	cMsg = "First Name is a required field." & vbCRLF
end if
If document.all.TxtNAME_LAST.value = "" then
	cMsg = cMsg & "Last Name is a required field." & vbCRLF
end if
If document.all.TxtLICENSE_NUMBER.value = "" then
	cMsg = cMsg & "License Number is a required field." & vbCRLF
end if
if len( cMsg ) <> 0 then
	MsgBox cMsg, 0, "FNSNetDesigner"
	ValidateScreenData = false
end if

End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.DID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.DID.value = "NEW"
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

	if not ValidateScreenData then 
		ExeSave = false
		exit function
	end if
	
	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.DID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		sResult = "Driver_ID"& Chr(129) & document.all.DID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "POLICY_ID" & Chr(129) & document.all.TxtPOLICY_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SSN"& Chr(129) & document.all.TxtSSN.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME_FIRST"& Chr(129) & document.all.TxtNAME_FIRST.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME_LAST"& Chr(129) & document.all.TxtNAME_LAST.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS1"& Chr(129) & document.all.TxtADDRESS1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS2"& Chr(129) & document.all.TxtADDRESS2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.CITY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.STATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.ZIP.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "RELATION_TO_INSURED" & Chr(129) & document.all.TxtRELATION_TO_INSURED.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LICENSE_NUMBER"& Chr(129) & document.all.TxtLICENSE_NUMBER.value & Chr(129) & "1" & Chr(128)
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

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"
	
	if document.all.ChkEdit.checked = true then
		SetScreenFieldsReadOnly false,"LABEL"
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScreenFieldsReadOnly true,"DISABLED"
		document.body.setAttribute "ScreenMode", "RO"
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"
end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Sub RefCountRpt_onclick()
	If document.all.SpanSharedCount.innerText > 0 Then
		If document.all.DID.value <> "" And document.all.DID.value <> "NEW" Then
			paramID = document.all.DID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedAttribute=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
	Else
		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
	End If	
End	Sub

Sub RefCountRpt_onmouseover()
	If document.all.SpanSharedCount.innerText > 0 Then
		document.all.RefCountRpt.style.cursor = "HAND"
	Else
		document.all.RefCountRpt.style.cursor = "DEFAULT"
	End If
End Sub

Sub BtnAttachVehicle_OnClick
	MODE = document.body.getAttribute("ScreenMode")
If MODE = "RW" Then
	VehicleObj.VehicleID = VehicleID
	strURL = "VehicleMaintenance.asp?SECURITYPRIV=FNSD_DRIVER&CONTAINERTYPE=MODAL&SEARCHONLY=TRUE"
	showModalDialog  strURL  ,VehicleObj ,"center=yes"
	If VehicleObj.VehicleID <> "" Then
			document.body.setAttribute "ScreenDirty", "YES"	
			document.all.TxtVEHICLE_ID.value = VehicleObj.VehicleID
	End If
End If
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Driver Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="DriverSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" name="SearchDID" value="<%=Request.QueryString("SearchDID")%>">
<input type="hidden" name="SearchPOLICY_ID" value="<%=Request.QueryString("SearchPOLICY_ID")%>">
<input type="hidden" name="SearchSSN" value="<%=Request.QueryString("SearchSSN")%>">
<input type="hidden" name="SearchNAME_FIRST" value="<%=Request.QueryString("SearchNAME_FIRST")%>">
<input type="hidden" name="SearchNAME_LAST" value="<%=Request.QueryString("SearchNAME_LAST")%>">
<input type="hidden" name="SearchADDRESS" value="<%=Request.QueryString("SearchADDRESS")%>">
<input type="hidden" name="SearchCITY" value="<%=Request.QueryString("SearchCITY")%>">
<input type="hidden" name="SearchSTATE" value="<%=Request.QueryString("SearchSTATE")%>">
<input type="hidden" name="SearchZIP" value="<%=Request.QueryString("SearchZIP")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="DID" value="<%=Request.QueryString("DID")%>">

<%	

Function TruncateRuleText(inText)
	if not IsNull(inText) then
		If Len(inText) < 40 Then
			TruncateRuleText = inText
		Else
			TruncateRuleText = Mid ( inText, 1, 40) & " ..."
		End If
	end if
End Function

Function TruncateLookupText(inText)
	if not IsNull(inText) then
		If Len(inText) < 22 Then
			TruncateLookupText = inText
		Else
			TruncateLookupText = Mid ( inText, 1, 22) & " ..."
		End If
	end if
End Function

Function ReplaceRuleText(inText)
	if not IsNull(inText) then
		ReplaceRuleText = Replace(inText,"""","&quot;")
	end if
End Function
If DID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td WIDTH="14"><img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count"></td><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">:<span id="SpanSharedCount"><%=SharedCount%></span></td>

<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL">
	<tr>
	<td COLSPAN="5" CLASS="LABEL">Driver ID:&nbsp;<span id="spanDID"><%=Request.QueryString("DID")%></span></td>
	</tr>
	<tr>
	<td VALIGN="BOTTOM"><img NAME="BtnAttachPolicy" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Policy">
	<td CLASS="LABEL" COLSPAN="2">Policy ID:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" READONLY STYLE="BACKGROUND-COLOR:SILVER" MAXLENGTH="10" TYPE="TEXT" NAME="TxtPOLICY_ID" VALUE="<%=RSPOLICY_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>	<td CLASS="LABEL">SSN:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="9" size="9" TYPE="TEXT" NAME="TxtSSN" VALUE="<%=RSSSN%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">First Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="40" size="40" TYPE="TEXT" NAME="TxtNAME_FIRST" VALUE="<%=RSNAME_FIRST%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Last Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" size="40" TYPE="TEXT" NAME="TxtNAME_LAST" VALUE="<%=RSNAME_LAST%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>
<table>
<tr>
	<td CLASS="LABEL">Address 1:<br><input ScrnInput="TRUE" size="98" TYPE="TEXT" MAXLENGTH="80" NAME="TxtADDRESS1" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%=RSADDRESS1%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Address 2:<br><input ScrnInput="TRUE" size="98" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtADDRESS2" VALUE="<%=RSADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
</table>
<table>
<tr>
	<td CLASS="LABEL">Zip:<br><input ScrnInput="TRUE" size="9" CLASS="LABEL" MAXLENGTH="9" TYPE="TEXT" NAME="Zip" VALUE="<%=RSZIP%>" ></td>	
	<td CLASS="LABEL">City:<br><input size="40" TYPE="TEXT" MAXLENGTH="40" NAME="CITY" CLASS="READONLY" READONLY TABINDEX=-1 VALUE="<%=RSCITY%>" ></td>
	<td CLASS="LABEL">State:<br><input size="3" TYPE="TEXT" MAXLENGTH="3" NAME="STATE" CLASS="READONLY" READONLY TABINDEX=-1></td>
</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Phone Number:<br><input ScrnInput="TRUE" size="15" TYPE="TEXT" MAXLENGTH="14" NAME="TxtPHONE" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSPHONE%>"></td>
	<td CLASS="LABEL">License Number:<br><input ScrnInput="TRUE" size="20" CLASS="LABEL" MAXLENGTH="20" TYPE="TEXT" NAME="TxtLICENSE_NUMBER" VALUE="<%=RSLICENSE_NUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Relation to Insured:<br><input ScrnInput="TRUE" size="40" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtRELATION_TO_INSURED" VALUE="<%=RSRELATION_TO_INSURED%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No driver selected.
</div>
<% End If %>

</form>
</body>
</html>


