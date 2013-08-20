<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%Response.Expires=0%>
<!--#include file="..\lib\ZIP.inc"-->
<%
AccountTextLen = 30	
Dim  EID
EID	= CStr(Request.QueryString("EID"))

RSPOLICY_ID = Request.QueryString("PID")
	
If EID <> "" Then
	If EID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT EMPLOYEE.*, ACCOUNT_HIERARCHY_STEP.NAME FROM EMPLOYEE, ACCOUNT_HIERARCHY_STEP WHERE " &_
				"EMPLOYEE.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				"EMPLOYEE_ID = " & EID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			ACCNT_HRCY_STEP_ID = RS("ACCNT_HRCY_STEP_ID")
			ACCOUNT_NAME = ReplaceQuotesInText(RS("NAME"))
			SSN = RS("SSN")
			NAME_FIRST = ReplaceQuotesInText(RS("NAME_FIRST"))
			NAME_LAST = ReplaceQuotesInText(RS("NAME_LAST"))
			TITLE = ReplaceQuotesInText(RS("TITLE"))
			ADDRESS1 = ReplaceQuotesInText(RS("ADDRESS1"))
			ADDRESS2 = ReplaceQuotesInText(RS("ADDRESS2"))
			CITY = RS("CITY")
			STATE = RS("STATE")
			ZIP = RS("ZIP")
			PHONE = RS("PHONE")
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
<title>Employee Details</title>
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
dim g_StatusInfoAvailable
g_StatusInfoAvailable =  false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "EmployeeSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateEID(inEID)
	document.all.EID.value = inEID
	document.all.spanEID.innerText = inEID
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

Function GetEID
	if document.all.EID.value <> "NEW" then
		GetEID = document.all.EID.value
	else
		GetEID = ""
	end if 
End Function

Function GetEIDName
	GetEIDName = document.all.TxtName.value
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

errstr = ""
	If document.all.TxtNAME_LAST.value = "" then
		errstr = errstr &  "Last Name is a required field." & VbCrlf
	end if
	If document.all.AHSID_ID.innerText <> "" AND NOT IsNumeric(document.all.AHSID_ID.innerText) then
		errstr = errstr &  "AHS ID must be numeric." & VbCrlf
	end if
	If document.all.AHSID_ID.innerText = "" then
		errstr = errstr &  "AHS ID is a required field."& VbCrlf
	end if
	If document.all.TxtSSN.value = "" then
		errstr = errstr &  "SSN is a required field."& VbCrlf
	end if
	if errstr = "" Then
		ValidateScreenData = true
	else
		msgbox errstr, 0 , "FNSDesigner"
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
	
	if document.all.EID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.EID.value = "NEW"
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
	
	if document.all.EID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.EID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "EMPLOYEE_ID"& Chr(129) & document.all.EID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "SSN"& Chr(129) & document.all.TxtSSN.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "NAME_LAST"& Chr(129) & document.all.TxtNAME_LAST.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME_FIRST"& Chr(129) & document.all.TxtNAME_FIRST.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TITLE"& Chr(129) & document.all.TxtTITLE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS1"& Chr(129) & document.all.TxtADDRESS1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS2"& Chr(129) & document.all.TxtADDRESS2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.CITY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.STATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.ZIP.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
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

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

'Sub RefCountRpt_onclick()
'	If document.all.SpanSharedCount.innerText > 0 Then
'		If document.all.EID.value <> "" And document.all.EID.value <> "NEW" Then
'			paramID = document.all.EID.value
'		Else	
'			paramID = 0
'		End If
'		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedAttribute=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
'	Else
'		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
'	End If	
'End	Sub
'Sub RefCountRpt_onmouseover()
'	If document.all.SpanSharedCount.innerText > 0 Then
'		document.all.RefCountRpt.style.cursor = "HAND"
'	Else
'		document.all.RefCountRpt.style.cursor = "DEFAULT"
'	End If
'End Sub

Function InEditMode
	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "This screen is read only.",0,"FNSNetDesigner"
		InEditMode = false
	End If
End Function


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
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_EMPLOYEE&SELECTONLY=TRUE&AHSID=" &AHSID
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Employee Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="EmployeeSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" name="SearchEID" value="<%=Request.QueryString("SearchEID")%>">
<input type="hidden" name="SearchSSN" value="<%=Request.QueryString("SearchSSN")%>">
<input type="hidden" name="SearchNAME_LAST" value="<%=Request.QueryString("SearchNAME_LAST")%>">
<input type="hidden" name="SearchNAME_FIRST" value="<%=Request.QueryString("SearchNAME_FIRST")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="EID" value="<%=Request.QueryString("EID")%>">

<%	
If EID <> "" Then

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
<table class="LABEL">
<tr>
	<td>
	<img NAME="BtnAttachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<img NAME="BtnDetachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	</td>
	<td width="305" nowrap>Account:&nbsp;<span ID="AHSID_TEXT" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(ACCOUNT_NAME)%>"><%=TruncateText(ACCOUNT_NAME,AccountTextLen)%></span></td>
	<td>A.H.Step ID:&nbsp;<span ID="AHSID_ID" CLASS="LABEL"><%=ACCNT_HRCY_STEP_ID%></span></td>
	</tr>
</table>

<span CLASS="LABEL">Employee ID:&nbsp;<span id="spanEID"><%=Request.QueryString("EID")%></span>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL" BORDER="0">
	<tr>
	<td CLASS="LABEL">SSN:<br><input ScrnInput="TRUE" size="28" CLASS="LABEL" MAXLENGTH="9" TYPE="TEXT" NAME="TxtSSN" VALUE="<%=SSN%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Phone:<br><input ScrnInput="TRUE" size="28" CLASS="LABEL" MAXLENGTH="14" TYPE="TEXT" NAME="TxtPHONE" VALUE="<%=PHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>

	<td CLASS="LABEL">Last Name:<br><input ScrnInput="TRUE" size="38" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtNAME_LAST" VALUE="<%=NAME_LAST%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">First Name:<br><input ScrnInput="TRUE" size="38" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" NAME="TxtNAME_FIRST" VALUE="<%=NAME_FIRST%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="3">Address 1:<br><input ScrnInput="TRUE" size="80" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtADDRESS1" VALUE="<%=ADDRESS1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="3">Address 2:<br><input ScrnInput="TRUE" size="80" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtADDRESS2" VALUE="<%=ADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="3">Title:<br><input ScrnInput="TRUE" size="80" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtTITLE" VALUE="<%=TITLE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL" ALIGN="LEFT">Zip:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="9" size="10" TYPE="TEXT" NAME="ZIP" VALUE="<%=ZIP%>" ></td>	
	<td CLASS="LABEL">City:<br><input CLASS="READONLY" READONLY TABINDEX=-1 MAXLENGTH="40" size="25" TYPE="TEXT" NAME="CITY" VALUE="<%=CITY%>" ></td>
	<td CLASS="LABEL">State:<br><input CLASS="READONLY" READONLY TABINDEX=-1 MAXLENGTH="3" size="3" TYPE="TEXT" NAME="STATE" VALUE="<%=STATE%>" ></td>
	</tr>
	</table>
	
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Employee selected.
</div>
<% End If %>
</form>
</body>
</html>


