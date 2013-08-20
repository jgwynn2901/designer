<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	Response.Expires=0 %>
<html>
<head>
<title>Coverage Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
var g_StatusInfoAvailable = false;
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<% if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<% end if %>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "CoverageSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateCOVID(inCOVID)
	document.all.COVID.value = inCOVID
	document.all.spanCOVID.innerText = inCOVID
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

Function GetCOVID
	if document.all.COVID.value <> "NEW" then
		GetCOVID = document.all.COVID.value
	else
		GetCOVID = ""
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
errmsg = ""

	if document.all.TxtDEDUCTIBLE.value <> "" then
		if Isnumeric(document.all.TxtDEDUCTIBLE.value) = false then
			errmsg = errmsg & "Deductible must be numeric." & VbCrlf
		end if
	end if
	if document.all.TxtLIMIT.value <> "" then
		if Isnumeric(document.all.TxtLIMIT.value) = false then
			errmsg = errmsg & "Limit must be numeric." & VbCrlf
		end if
	end if
	if document.all.TxtPOLICY_ID.value <> "" then
		if Isnumeric(document.all.TxtPOLICY_ID.value) = false then
			errmsg = errmsg & "Policy Id must be numeric." & VbCrlf
		end if
	end if
	if document.all.TxtVEHICLE_ID.value <> "" then
		if Isnumeric(document.all.TxtVEHICLE_ID.value) = false then
			errmsg = errmsg & "Vehicle Id must be numeric." & VbCrlf
		end if
	end if
	if document.all.TxtEFFECTIVE_DATE.value <> "" then
		if CheckDate(document.all.TxtEFFECTIVE_DATE.value) = false then
			errmsg = errmsg & "Effective Date formatted incorrectly.  (MM/DD/YYYY)" & VbCrlf
		end if
	end if
	if document.all.TxtEXPIRATION_DATE.value <> "" then
		if CheckDate(document.all.TxtEXPIRATION_DATE.value) = false then
			errmsg = errmsg & "Expiration Date is formatted incorrectly. (MM/DD/YYYY)" & VbCrlf
		end if
	end if
	if document.all.TxtTYPE.value = "" then
			errmsg = errmsg & "Type is a required field." & VbCrlf
	end if
	If errmsg = "" Then
		ValidateScreenData = true		
	Else
		msgbox errmsg, 0, "FNSDesigner"
		ValidateScreenData = false
	End If
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.COVID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.COVID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave()
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.COVID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.COVID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "COVERAGE_ID"& Chr(129) & document.all.COVID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "POLICY_ID"& Chr(129) & document.all.TxtPOLICY_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "VEHICLE_ID"& Chr(129) & document.all.TxtVEHICLE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TYPE"& Chr(129) & document.all.TxtTYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DEDUCTIBLE"& Chr(129) & document.all.TxtDEDUCTIBLE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LIMIT"& Chr(129) & document.all.TxtLIMIT.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "EFFECTIVE_DATE" & Chr(129) & "TO_DATE('" & document.all.TxtEFFECTIVE_DATE.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		sResult = sResult & "EXPIRATION_DATE" & Chr(129) &  "TO_DATE('" & document.all.TxtEXPIRATION_DATE.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	'Else
	''	SpanStatus.innerHTML = "Nothing to Save"
	'End If
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

Function CheckDate( InDate )
If IsNumeric(Mid(InDate,1,2)) AND IsNumeric(Mid(InDate,4,2)) Then
	If Not IsDate(InDate) Then
		CheckDate = false
		Exit Function
	End If
	If Len(InDate) <> 10 OR Mid(InDate,1,2) > 12 Then
		CheckDate = false
		Exit Function
	End If
	CheckDate = true
Else
	CheckDate = true
End If
End Function
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Coverage Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="CoverageSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchCOVID" value="<%=Request.QueryString("SearchCOVID")%>">
<input type="hidden" name="SearchPOLICY_ID" value="<%=Request.QueryString("SearchPOLICY_ID")%>">
<input type="hidden" name="SearchVEHICLE_ID" value="<%=Request.QueryString("SearchVEHICLE_ID")%>">
<input type="hidden" name="SearchEFFECTIVE_DATE" value="<%=Request.QueryString("SearchEFFECTIVE_DATE")%>">
<input type="hidden" name="SearchEXPIRATION_DATE" value="<%=Request.QueryString("SearchEXPIRATION_DATE")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="COVID" value="<%=Request.QueryString("COVID")%>">

<%	
Dim COVID
COVID	= CStr(Request.QueryString("COVID"))

If CStr(Request.QueryString("VID")) <> "" Then 
	RSVEHICLE_ID = CStr(Request.QueryString("VID"))
ElseIf	CStr(Request.QueryString("PID")) <> "" Then 
	RSPOLICY_ID = CStr(Request.QueryString("PID"))
End If
	

If COVID <> "" Then
	If COVID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = ""
		SQLST = SQLST & "SELECT COVERAGE_ID, POLICY_ID, VEHICLE_ID, TYPE "
		SQLST = SQLST & ", DEDUCTIBLE,LIMIT, TO_CHAR(EFFECTIVE_DATE, 'MM/DD/YYYY') As EFFECTIVE_DATE , TO_CHAR(EXPIRATION_DATE, 'MM/DD/YYYY') As EXPIRATION_DATE FROM COVERAGE WHERE COVERAGE_ID = " & COVID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
		
			RSCOVERAGE_ID = RS("COVERAGE_ID")
			RSPOLICY_ID = RS("POLICY_ID")
			RSVEHICLE_ID = RS("VEHICLE_ID")
			RSTYPE = ReplaceQuotesInText(RS("TYPE"))
			RSDEDUCTIBLE = RS("DEDUCTIBLE")
			RSLIMIT = RS("LIMIT")
			RSEFFECTIVE_DATE = RS("EFFECTIVE_DATE")
			RSEXPIRATION_DATE = RS("EXPIRATION_DATE")
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
	
%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<span CLASS="LABEL">Coverage ID:&nbsp;<span id="spanCOVID"><%=Request.QueryString("COVID")%></span></span>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr><td>
<table class="LABEL">
<tr>
	<td CLASS="LABEL">Policy Id:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" SIZE="10" TYPE="TEXT" NAME="TxtPOLICY_ID" VALUE="<%=RSPOLICY_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Vehicle Id:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtVEHICLE_ID" VALUE="<%=RSVEHICLE_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Type:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" size="20" TYPE="TEXT" NAME="TxtTYPE" VALUE="<%=RSTYPE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Deductible:<br><input type="text" Size="12" MAXLENGTH="12" Scrninput="TRUE" NAME="TxtDEDUCTIBLE" VALUE="<%= RSDEDUCTIBLE %>" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Limit:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="12" size="12" TYPE="TEXT" NAME="TxtLimit" VALUE="<%=RSLIMIT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Effective Date:<br><input ScrnInput="TRUE" size="12" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtEFFECTIVE_DATE" VALUE="<%=RSEFFECTIVE_DATE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Expiration Date:<br><input ScrnInput="TRUE" CLASS="LABEL" size="12" MAXLENGTH="10" TYPE="TEXT" NAME="TxtEXPIRATION_DATE" VALUE="<%= RSEXPIRATION_DATE %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
 <% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No coverage selected.
</div>
<% End If %>
</form>
</body>
</html>


