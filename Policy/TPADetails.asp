<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<!--#include file="..\lib\CheckSharedCarrier.inc"-->
<%
Response.Expires=0 %>
<!--#include file="..\lib\ZIP.inc"-->
<%
	Dim SharedCount, SharedCountText, TPAID
	SharedCount = 0
	SharedCountText = "Ready"
	
	TPAID	= CStr(Request.QueryString("TPAID"))

	If TPAID <> "" Then
		If TPAID = "NEW" Then 
			SharedCount = 0
		Else
			'SharedCount = CheckSharedTPA(CLng(TPAID),True,True,1,False,False,0)
		End If
	End If	
	
	
If TPAID <> "" Then
	If TPAID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM THIRD_PARTY_ADMINISTRATOR WHERE TPA_ID = " & TPAID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSNAME = RS("NAME")
			RSTITLE = RS("TITLE")
			RSBUSINESSTYPE = RS("BUSINESS_TYPE")
			RSADDRESS1 = RS("ADDRESS1")
			RSADDRESS2 = RS("ADDRESS2")
			RSCITY = RS("CITY")
			RSSTATE = RS("STATE")
			RSZIP = RS("ZIP")
			RSPHONE = RS("PHONE")
			RSFEIN	= RS("FEIN")
			RSTPANUMBER = RS("TPA_NUMBER")
			RSBUREAUCD = RS("BUREAU_CD")
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
<title>TPA Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if TPAID <> "" then %>
			document.all.State.Value = "<%= RSSTATE %>"
			<% if SharedCount <= 1 then %>
			document.all.ChkEdit.checked = true
			ChkEdit_OnClick
			<%	else %>
					SetStatusInfoAvailableFlag(true)
					document.all.ChkEdit.checked = false
					ChkEdit_OnClick
				<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
					If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
						document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 Color='Maroon'>+</Font>"
				<%	Else %>
						document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>"
				<%  End If
				end if
		end if	
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "TPASearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateTPAID(inTPAID)
	document.all.TPAID.value = inTPAID
	document.all.spanTPAID.innerText = inTPAID
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

Function GetTPAID
	if document.all.TPAID.value <> "NEW" then
		GetTPAID = document.all.TPAID.value
	else
		GetTPAID = ""
	end if 
End Function

Function GetTPAIDName
	GetTPAIDName = document.all.TxtName.value
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
	If  document.all.TxtName.value = "" then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
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
	
	if document.all.TPAID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.TPAID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	document.all.SpanSharedCount.innerText = 0
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
	
	if document.all.TPAID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.TPAID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "TPA_ID"& Chr(129) & document.all.TPAID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TITLE"& Chr(129) & document.all.TxtTitle.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BUSINESS_TYPE"& Chr(129) & document.all.TxtBUSINESSTYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS1"& Chr(129) & document.all.TxtADDRESS1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS2"& Chr(129) & document.all.TxtADDRESS2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.CITY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.STATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.ZIP.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FEIN"& Chr(129) & document.all.TxtFEIN.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TPA_NUMBER"& Chr(129) & document.all.TxtTPANumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BUREAU_CD"& Chr(129) & document.all.TxtBureauCD.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"

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
		If document.all.TPAID.value <> "" And document.all.TPAID.value <> "NEW" Then
			paramID = document.all.TPAID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedTPA=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
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
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» TPA Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="TPASave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchTPAID" value="<%=Request.QueryString("SearchTPAID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchTitle" value="<%=Request.QueryString("SearchTitle")%>">
<input type="hidden" name="SearchBusinessType" value="<%=Request.QueryString("SearchBusinessType")%>">
<input type="hidden" name="SearchAddress" value="<%=Request.QueryString("SearchAddress")%>">
<input type="hidden" name="SearchCity" value="<%=Request.QueryString("SearchCity")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" name="SearchTPANumber" value="<%=Request.QueryString("SearchTPANumber")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="TPAID" value="<%=Request.QueryString("TPAID")%>">
<input type="hidden" NAME="BUREAUCD" value="<%=Request.QueryString("BUREAUCD")%>">
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
If TPAID <> "" Then
%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td WIDTH="14">
<img ID="RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count">
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">
:<span id="SpanSharedCount"><%=SharedCount%></span>
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
<input ScrnBtn="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL">
	<tr>
	<td CLASS="LABEL">TPA ID:&nbsp;<span id="spanTPAID"><%=Request.QueryString("TPAID")%></span></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="255" SIZE="85" TYPE="TEXT" NAME="TxtName" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
		<table>
		<tr>
			<td CLASS="LABEL">TPA Number:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" size="30" TYPE="TEXT" NAME="TxtTPANumber" VALUE="<%=RSTPANUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
			<td CLASS="LABEL">Bureau Code:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="15" size="15" TYPE="TEXT" NAME="TxtBureauCD" VALUE="<%=RSBUREAUCD%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
		</tr>
		<table>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Title:<br><input ScrnInput="TRUE" size="85" CLASS="LABEL" MAXLENGTH="85" TYPE="TEXT" NAME="TxtTitle" VALUE="<%=RSTITLE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Business Type:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" size="30" TYPE="TEXT" NAME="TxtBusinessType" VALUE="<%=RSBUSINESSTYPE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Phone:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="20" size="20" TYPE="TEXT" NAME="TxtPhone" VALUE="<%=RSPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Fein:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="20" size="20" TYPE="TEXT" NAME="TxtFein" VALUE="<%=RSFEIN%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Address 1:<br><input ScrnInput="TRUE" size="86" TYPE="TEXT" MAXLENGTH="45" NAME="TxtAddress1" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSADDRESS1%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Address 2:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="45" size="86" TYPE="TEXT" NAME="TxtAddress2" VALUE="<%=RSADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Zip:<br><input ScrnInput="TRUE" size="9" CLASS="LABEL" MAXLENGTH="9" TYPE="TEXT" NAME="Zip" VALUE="<%=RSZIP%>" ></td>	
	<td CLASS="LABEL">City:<br><input size="30" CLASS="READONLY" READONLY MAXLENGTH="30" TYPE="TEXT" NAME="City" TABINDEX=-1 VALUE="<%=RSCITY%>" ></td>
	<td CLASS="LABEL">State:<br><input size="3" CLASS="READONLY" READONLY MAXLENGTH="3" TYPE="TEXT" NAME="STATE" TABINDEX=-1 VALUE="<%=RSSTATE%>" ></td>
</tr>
</table>
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No TPA selected.
</div>
<% End If %>
</form>
</body>
</html>