<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
	Response.Expires=0 
	AccountTextLen = 30	

	Dim SharedCount, SharedCountText, cFeeID
	dim cSQL, oRS, oConn, cAHSID
	
	SharedCount = 0
	SharedCountText = "Ready"
	
	cFeeID = trim(Request.QueryString("BID"))
	cAHSID = trim(Request.QueryString("AHSID"))
	If cFeeID <> "" Then
		If cFeeID = "NEW" Then 
			SharedCount = 0
		'Else
		'	SharedCount = CheckSharedAttribute(CLng(BID),True,True,1,False,False,0)
		End If
	End If	
	
If len(cFeeID) <> 0 Then
	If cFeeID <> "NEW" then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		cSQL = "SELECT FEE.*, ACCOUNT_HIERARCHY_STEP.NAME FROM FEE, ACCOUNT_HIERARCHY_STEP " &_
				"WHERE FEE.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				"FEE_ID = " & cFeeID
		Set oRS = oConn.Execute(cSQL)
		If Not oRS.EOF then
			RSFEE_ID = oRS("FEE_ID")
			RSFEE_TYPE_ID = oRS("FEE_TYPE_ID")
			RSACCNT_HRCY_STEP_ID= oRS("ACCNT_HRCY_STEP_ID")
			RSACCOUNT_NAME= ReplaceQuotesInText(oRS("NAME"))
			RSLOB_CD = oRS("LOB_CD")
			RSCALL_TYPE = oRS("CALL_TYPE")
			RSFEE_AMOUNT = oRS("FEE_AMOUNT")
			RSFREE_PERCENTAGE = oRS("FREE_PERCENTAGE")
			RSBEGIN_CALL_RANGE = oRS("BEGIN_CALL_RANGE")
			RSEND_CALL_RANGE = oRS("END_CALL_RANGE")
			RSFREE_COUNT = oRS("FREE_COUNT")
			RSREASON_CODE = oRS("REASON_CODE")
			RSDESCRIPTION = ReplaceQuotesInText(oRS("DESCRIPTION"))
		end if	
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	else
		if len(cAHSID) <> 0 then
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.Open CONNECT_STRING
			cSQL = "SELECT ACCOUNT_HIERARCHY_STEP.NAME FROM ACCOUNT_HIERARCHY_STEP " &_
					"WHERE ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID = " & Request.QueryString("AHSID")
			Set oRS = oConn.Execute(cSQL)
			If Not oRS.EOF then
				RSACCNT_HRCY_STEP_ID= cAHSID
				RSACCOUNT_NAME= ReplaceQuotesInText(oRS("NAME"))
			end if
			oRS.close
			oConn.close
			set oRS = nothing
			set oConn = nothing	
		end if
		RSBEGIN_CALL_RANGE = "0"
		RSEND_CALL_RANGE = "0"
		RSFREE_COUNT = "0"
		RSFREE_PERCENTAGE = "0"
	end if	
End If
%>
<html>
<head>
<title>Fee Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}

var AHSSearchObj = new CAHSSearchObj();
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
dim cLOB
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if len(cFeeID) <> 0 and cFeeID<>"NEW" then %>
		cLOB = "<%= RSLOB_CD%>"
		document.all.TxtLOB_CD.VAlue = cLOB
		document.all.TxtCallType.value = "<%=RSCALL_TYPE%>"
		document.all.TxtFEE_TYPE_ID.VAlue = "<%= RSFEE_TYPE_ID%>"
		<%
		if SharedCount <= 1 then %>
<%	else %>
	SetStatusInfoAvailableFlag(true)
<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "BillingSearch-f.asp"
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

Function GetBIDName
	GetBIDName = document.all.TxtName.value
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
dim cErrmsg

If document.all.AHSID_ID.innerText = "" then
	cErrmsg = "Accnt Hrcy Step ID is a required field." & VbCrlf
end if
If document.all.TxtLOB_CD.value = "" then
	cErrmsg = cErrmsg & "Line of Business is a required field." & VbCrlf
'elseif document.all.TxtLOB_CD.value = "INF" then
'	if document.all.TxtCallType.value <> "I" then
'		cErrmsg = cErrmsg & "For LOB 'INF' the Call Type must be 'I'." & VbCrlf
'	end if
end if

If document.all.TxtFEE_TYPE_ID.value = "" then
	cErrmsg = cErrmsg & "Fee Type is a required field." & VbCrlf
end if
If document.all.TxtFREE_PERCENTAGE.value = "" then
	cErrmsg = cErrmsg & "Free Percentage is a required field." & VbCrlf
elseIf Not Isnumeric(document.all.TxtFREE_PERCENTAGE.value) then
	cErrmsg = cErrmsg & "Free Percentage must be numeric." & VbCrlf
end if
If document.all.TxtFEE_AMOUNT.value = "" then
	cErrmsg = cErrmsg & "Fee Amount is a required field." & VbCrlf
elseif Not IsNumeric(document.all.TxtFEE_AMOUNT.value) then
	cErrmsg = cErrmsg & "Fee Amount must be numeric." & VbCrlf
	ValidateScreenData = false
end if
If document.all.TxtBEGIN_CALL_RANGE.value <> "" AND Not IsNumeric(document.all.TxtBEGIN_CALL_RANGE.value) then
	cErrmsg = cErrmsg & "Begin Call Range must be numeric." & VbCrlf
	ValidateScreenData = false
end if
If document.all.TxtEND_CALL_RANGE.value <> "" AND Not IsNumeric(document.all.TxtEND_CALL_RANGE.value) then
	cErrmsg = cErrmsg & "End Call Range must be numeric." & VbCrlf
	ValidateScreenData = false
end if
If document.all.TxtFREE_COUNT.value <> "" AND Not IsNumeric(document.all.TxtFREE_COUNT.value) then
	cErrmsg = cErrmsg & "Free Count must be numeric." & VbCrlf
	ValidateScreenData = false
end if
'If document.all.TxtBEGIN_CALL_RANGE.value <> "" AND document.all.TxtREASON_CODE.value = "" then
'	cErrmsg = cErrmsg & "Reason Code is a required field." & VbCrlf
'	ValidateScreenData = false
'end if

If len(cErrmsg) = 0 Then
	ValidateScreenData = true
Else
	msgbox cErrmsg, 0, "FNSDesigner"
End If
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
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.BID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "FEE_ID"& Chr(129) & document.all.BID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "0" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOB_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FEE_TYPE_ID"& Chr(129) & document.all.TxtFEE_TYPE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CALL_TYPE"& Chr(129) & document.all.TxtCALLTYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FEE_AMOUNT"& Chr(129) & document.all.TxtFEE_AMOUNT.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FREE_PERCENTAGE"& Chr(129) & document.all.TxtFREE_PERCENTAGE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BEGIN_CALL_RANGE"& Chr(129) & document.all.TxtBEGIN_CALL_RANGE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "END_CALL_RANGE"& Chr(129) & document.all.TxtEND_CALL_RANGE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FREE_COUNT"& Chr(129) & document.all.TxtFREE_COUNT.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "REASON_CODE"& Chr(129) & document.all.TxtREASON_CODE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDESCRIPTION.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If
	ExeSave = bRet
End Function

sub LOB_OnChange(cLOB)
dim lDisabled

lDisabled = false
if cLOB <> "INF" then
	lDisabled = true
end if
document.all.TxtFREE_PERCENTAGE.disabled = lDisabled
document.all.TxtFreeInfoClaims.disabled = lDisabled
end sub

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

'Sub RefCountRpt_onclick()
''	If document.all.SpanSharedCount.innerText > 0 Then
'		If document.all.BID.value <> "" And document.all.BID.value <> "NEW" Then
'			paramID = document.all.BID.value
'		Else	
'			paramID = 0
'		End If
'		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedAttribute=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
'	Else
'		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
'	End If	
'End Sub

'Sub RefCountRpt_onmouseover()
'	If document.all.SpanSharedCount.innerText > 0 Then
'		document.all.RefCountRpt.style.cursor = "HAND"
'	Else
'		document.all.RefCountRpt.style.cursor = "DEFAULT"
'	End If
'End Sub

Function AttachAccount (ID, SPANID)
	AHSID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID.title
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_FEE&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,AHSSearchObj ,"center"

	'if Selected=true update everything, otherwise if AHSID is the same, update text in case of save
	If AHSSearchObj.Selected = true Then
		If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = AHSSearchObj.AHSID
		end if
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Fee Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<form Name="FrmDetails" METHOD="POST" ACTION="BillingSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" name="SearchBID" value="<%=Request.QueryString("SearchBID")%>">
<input type="hidden" name="SearchACCNT_HRCY_STEP_ID" value="<%=Request.QueryString("SearchACCNT_HRCY_STEP_ID")%>">
<input type="hidden" name="SearchLOB_CD" value="<%=Request.QueryString("SearchLOB_CD")%>">
<input type="hidden" name="SearchFEE_TYPE_ID" value="<%=Request.QueryString("SearchFEE_TYPE_ID")%>">
<input type="hidden" name="SearchDESCRIPTION" value="<%=Request.QueryString("SearchDESCRIPTION")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="BID" value="<%=Request.QueryString("BID")%>">

<%	

If cFeeID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<!--<td WIDTH="14"><img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count"></td><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">:<span id="SpanSharedCount"><%=SharedCount%></span></td>-->
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
<input ScrnBtn="TRUE" STYLE="DISPLAY:NONE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL">
<tr>
	<td>
	<img NAME="BtnAttachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<img NAME="BtnDetachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	</td>
	<td width="305" nowrap>Account:&nbsp;<span ID="AHSID_TEXT" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(RSACCOUNT_NAME)%>"><%=TruncateText(RSACCOUNT_NAME,AccountTextLen)%></span></td>
	<td>A.H.Step ID:&nbsp;<span ID="AHSID_ID" CLASS="LABEL"><%=RSACCNT_HRCY_STEP_ID%></span></td>
	</tr>
</table>

<table CLASS="LABEL" CELLPADDING="3" CELLSPACING="0" border="0" ID="Table1">
<tr>
<td COLSPAN="4">Billing ID:&nbsp<span id="spanBID"><%=Request.QueryString("BID")%></span></td>
</tr>
	<tr>
	<td CLASS="LABEL"><nobr>Fee Type:<br>
	<select NAME="TxtFEE_TYPE_ID" CLASS="LABEL" ScrnBtn="TRUE" ID="Select1">
	<%
	SQLST = "SELECT * FROM FEE_TYPE WHERE NAME IS NOT NULL"
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	Set RS2 = Conn.Execute(SQLST)
	Do While Not RS2.EOF
		Response.Write "<option VALUE=" & RS2("FEE_TYPE_ID").value
		if UCase(RS2("NAME")) = "SERVICE FEE" then
			Response.Write " SELECTED"
		end if
		Response.Write ">" & RS2("NAME").value & chr(10) & chr(13)
		RS2.MoveNext
	Loop
	RS2.CLose
	%>
	</select></td>
	<td CLASS="LABEL">LOB:<br>
	<!--<select NAME="TxtLOB_CD" CLASS="LABEL" ScrnBtn="TRUE" onchange="chkINF(this.options[this.selectedIndex].value)">-->
	<select language="jscript" NAME="TxtLOB_CD" CLASS="LABEL" ScrnBtn="TRUE" ID="Select2" ONCHANGE="LOB_OnChange(this.options[this.selectedIndex].value)">
	<option VALUE>
	<%
	SQLST = "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
	Set RS = Conn.Execute(SQLST)
	Do While Not RS.EOF
	%>
	<option VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_CD") %>
	<%
	RS.MoveNext
	Loop
	RS.CLose
	Conn.close
	set Conn = nothing
	%>
	</select></td>
	<td CLASS="LABEL" >Call Type:<br>
	<select ID="TxtCallType" CLASS="LABEL" ScrnBtn="TRUE" NAME="TxtCallType">
		<option VALUE="I"> Info</option>
		<option VALUE="C"> Call</option> 
		<option VALUE="F"> Fax</option>
		<option VALUE="N"> Net</option>
		<option VALUE="E"> Email</option>
		<option VALUE="T"> Trans </option>
		<option VALUE="O"> Ofcall </option>
		<option VALUE="W"> Web </option>
	</select></td>
	<td CLASS="LABEL">Service Fee Amount:<br>$ <input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFEE_AMOUNT" VALUE="<%=RSFEE_AMOUNT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text1"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="3">Description:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text2"></td>
	<td><br></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Reason Code:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" size="50" TYPE="TEXT" NAME="TxtREASON_CODE" VALUE="<%=RSREASON_CODE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text3"></td>
	<td COLSPAN="2"><br></td>
	</tr>
	<tr>
	<td CLASS="LABEL" width="120px" rowspan="2" style="border-style: solid">Free Output Pages:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFREE_COUNT" VALUE="<%=RSFREE_COUNT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text4"></td>
	<td CLASS="LABEL" rowspan="2" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-bottom-style: solid"><b>Tiered Billing</b></td>
	<td COLSPAN="2" style="border-right-style: solid; border-top-style: solid; border-bottom-style: solid; border-bottom-width: 1">Begin Call Range:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtBEGIN_CALL_RANGE" VALUE="<%=RSBEGIN_CALL_RANGE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text5"></td>
	</tr>

	<tr>
	<td COLSPAN="2" style="border-right-style: solid; border-bottom-style: solid">End Call Range:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtEND_CALL_RANGE" VALUE="<%=RSEND_CALL_RANGE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text6"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2" rowspan="3" style="border-left-style: solid; border-right-style: solid; border-right-width: 1; border-bottom-style: solid" align="center" ><b>For Informational Calls</b></td>
	<td CLASS="LABEL" COLSPAN="2" style="border-right-style: solid" >Free Claims:<br><input DISABLED ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFreeInfoClaims" VALUE="0" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text7"></td>	
	</tr>
<tr>
<td COLSPAN="2" valign="top" align="left" style="border-right-style: solid"><b>OR</b></td>
</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2" style="border-right-style: solid; border-bottom-style: solid" >Free Percentage of Total Claims:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFREE_PERCENTAGE" VALUE="<%=RSFREE_PERCENTAGE %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text8">%</td>

	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="4" style="border-left-style: solid; border-right-style: solid; border-bottom-style: solid" >Minimum
      Quantity of Claims to Invoice:<br><input DISABLED ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtMinQuantity" VALUE="0" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text9">
      (0 = No Minimum)</td>

	</tr>

	</table>

<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No fee selected.
</div>
<% End If %>
</form>
</body>
</html>


