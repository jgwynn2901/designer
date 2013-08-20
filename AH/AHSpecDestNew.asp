<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
	Response.Expires=0 
	AccountTextLen = 30	

	Dim SharedCount, SharedCountText, BID
	SharedCount = 0
	SharedCountText = "Ready"
	
	BID	= CStr(Request.QueryString("BID"))

	If BID <> "" Then
		If BID = "NEW" Then 
			SharedCount = 0
		'Else
		'	SharedCount = CheckSharedAttribute(CLng(BID),True,True,1,False,False,0)
		End If
	End If	
	
If BID <> "" Then
RSACCNT_HRCY_STEP_ID = Request.QueryString("AHSID")
	If BID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT FEE.*, ACCOUNT_HIERARCHY_STEP.NAME FROM FEE, ACCOUNT_HIERARCHY_STEP " &_
				"WHERE FEE.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				"FEE_ID = " & BID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSFEE_ID = RS("FEE_ID")
			RSFEE_TYPE_ID = RS("FEE_TYPE_ID")
			RSACCNT_HRCY_STEP_ID= RS("ACCNT_HRCY_STEP_ID")
			RSACCOUNT_NAME= ReplaceQuotesInText(RS("NAME"))
			RSLOB_CD = RS("LOB_CD")
			RSCALL_TYPE = RS("CALL_TYPE")
			RSFEE_AMOUNT = RS("FEE_AMOUNT")
			RSFREE_PERCENTAGE = RS("FREE_PERCENTAGE")
			RSBEGIN_CALL_RANGE = RS("BEGIN_CALL_RANGE")
			RSEND_CALL_RANGE = RS("END_CALL_RANGE")
			RSFREE_COUNT = RS("FREE_COUNT")
			RSREASON_CODE = RS("REASON_CODE")
			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
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
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if BID <> "" then %>
		document.all.TxtLOB_CD.VAlue = "<%= RSLOB_CD%>"
		document.all.TxtFEE_TYPE_ID.VAlue = "<%= RSFEE_TYPE_ID%>"
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
errmsg = ""
	If document.all.AHSID_ID.innerText = "" then
		errmsg = errmsg & "Accnt Hrcy Step ID is a required field." & VbCrlf
	end if
	If document.all.TxtFEE_TYPE_ID.value = "" then
		errmsg = errmsg & "Fee Type is a required field." & VbCrlf
	end if
	If document.all.TxtFREE_PERCENTAGE.value = "" then
		errmsg = errmsg & "Free Percentage is a required field." & VbCrlf
	end if
		If document.all.TxtFREE_PERCENTAGE.value <> "" AND Not Isnumeric(document.all.TxtFREE_PERCENTAGE.value) then
		errmsg = errmsg & "Free Percentage must be numeric." & VbCrlf
	end if
	If document.all.TxtFEE_AMOUNT.value <> "" AND Not IsNumeric(document.all.TxtFEE_AMOUNT.value) then
		errmsg = errmsg & "Fee Amount must be numeric." & VbCrlf
		ValidateScreenData = false
	end if
	If document.all.TxtBEGIN_CALL_RANGE.value <> "" AND Not IsNumeric(document.all.TxtBEGIN_CALL_RANGE.value) then
		errmsg = errmsg & "Begin Call Range must be numeric." & VbCrlf
		ValidateScreenData = false
	end if
	If document.all.TxtEND_CALL_RANGE.value <> "" AND Not IsNumeric(document.all.TxtEND_CALL_RANGE.value) then
		errmsg = errmsg & "End Call Range must be numeric." & VbCrlf
		ValidateScreenData = false
	end if
	If document.all.TxtFREE_COUNT.value <> "" AND Not IsNumeric(document.all.TxtFREE_COUNT.value) then
		errmsg = errmsg & "Free Count must be numeric." & VbCrlf
		ValidateScreenData = false
	end if
	If errmsg = "" Then
		ValidateScreenData = true
	Else
		msgbox errmsg, 0, "FNSDesigner"
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
		sResult = sResult & "CALL_TYPE"& Chr(129) & document.all.TxtCALL_TYPE.value & Chr(129) & "1" & Chr(128)
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

If BID <> "" Then

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

<table class="LABEL">
	<tr>
	<td COLSPAN="5" CLASS="LABEL">First name:&nbsp;<span id="spanBID"><%=Request.QueryString("BID")%></span></td>
	</tr>
	<tr>
	<td CLASS="LABEL">First Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" size="30" TYPE="TEXT" NAME="TxtREASON_CODE" VALUE="<%=RSREASON_CODE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Middle Initial::<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="3" size="3" TYPE="TEXT" NAME="TxtREASON_CODE" VALUE="<%=RSREASON_CODE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Last Name:<br><input ScrnInput="TRUE" size="30" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Title:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Address 1:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Address 2:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">City:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">State:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">ZIP Code:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Phone No.:<br><input ScrnInput="TRUE" size="76" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	
	</table>
	<table>
	<tr><td CLASS="LABEL">Free Count:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFREE_COUNT" VALUE="<%=RSFREE_COUNT %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL" COLSPAN="2">Call Type:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" size="30" TYPE="TEXT" NAME="TxtCALL_TYPE" VALUE="<%=RSCALL_TYPE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL" COLSPAN="2">LOB:<br>
	<select NAME="TxtLOB_CD" CLASS="LABEL" ScrnBtn="TRUE">
	<option VALUE>
	<%
		Set Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = CONNECT_STRING
		Conn.Open ConnectionString
		SQLST = ""
		SQLST = SQLST & "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
		Set RS = Conn.Execute(SQLST)
	Do While Not RS.EOF
	%>
	<option VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_CD") %>
	<%
	RS.MoveNext
	Loop
	RS.CLose
	%>
	</select></td>
	<td CLASS="LABEL"><nobr>Fee Type:<br>
	<select NAME="TxtFEE_TYPE_ID" CLASS="LABEL" ScrnBtn="TRUE">
	<option VALUE>
	<%
		SQLST = ""
		SQLST = SQLST & "SELECT * FROM FEE_TYPE WHERE NAME IS NOT NULL"
		Set RS2 = Conn.Execute(SQLST)
	Do While Not RS2.EOF 
	%>
	<option VALUE="<%= RS2("FEE_TYPE_ID") %>"><%= RS2("NAME") %>
	<%
	RS2.MoveNext
	Loop
	RS2.CLose
	%>
	</select></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Fee Amount:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFEE_AMOUNT" VALUE="<%=RSFEE_AMOUNT %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Free Percentage:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtFREE_PERCENTAGE" VALUE="<%=RSFREE_PERCENTAGE %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Begin Call Range:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtBEGIN_CALL_RANGE" VALUE="<%=RSBEGIN_CALL_RANGE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">End Call Range:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtEND_CALL_RANGE" VALUE="<%=RSEND_CALL_RANGE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
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


