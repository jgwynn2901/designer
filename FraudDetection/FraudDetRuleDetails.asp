<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim FDTID, FDRID, isRequired
	
	FDTID =  CStr(Request.QueryString("FDTID"))
	FDRID =  CStr(Request.QueryString("FDRID"))
	
	s_DisplayMsg = "Ready"

	BranchTextLen = 30
	RuleTextLen = 30
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Fraud Detection Rule Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
window.returnValue = false;

function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
var RuleSearchObj = new CRuleSearchObj();

function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}

</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

<!--#include file="..\lib\Help.asp"-->

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
End Sub

Function AttachRule (ID, SPANID, strTITLE)
	RID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
	
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
		
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_COVERAGE_CODE_XREF&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = RuleSearchObj.RID
		end if
		UpdateRuleText(SPANID)
	ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateRuleText(SPANID)
	End If

End Function


Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateRuleText (SPANID)
	If Len(RuleSearchObj.RIDText) < <%=RuleTextLen%> Then
		SPANID.innertext = RuleSearchObj.RIDText
	Else
		SPANID.innertext = Mid ( RuleSearchObj.RIDText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = RuleSearchObj.RIDText
End Sub


Sub UpdateFDRID(inFDRID)
	document.all.FDRID.value = inFDRID
	document.all.spanFDRID.innerText = inFDRID
	if document.all.spanFDRID.innerText <> "NEW" then
			document.body.setAttribute "IsThisRequired", "N"
	End If
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

Function GetFDRID
	if document.all.FDRID.value <> "NEW" then
	    GetFDRID = document.all.FDRID.value
	else
		GetFDRID = ""
	end if 
End Function

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Function f_CheckIsThisRequired
	If CStr(document.body.getAttribute("IsThisRequired")) = "Y" Then
		f_CheckIsThisRequired = true
	Else
		f_CheckIsThisRequired = False
	End if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	If document.all.TxtName.value = "" Then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	
	If document.all.TxtScore.value <> "" Then
		If not IsNumeric(document.all.TxtScore.value) then
			MsgBox "Please enter a number in the Score field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	Else
		MsgBox "Score is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	If document.all.RULE_ID.innerText = "" Then
		MsgBox "Rule is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if

	ValidateScreenData = true
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If
	
	If document.all.FDTID.value = "" Then
		ExeSave = false
		exit function
	End If
		
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.FDRID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "FRAUD_DETECTION_RULE_ID " & Chr(129) & document.all.FDRID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FRAUD_DETECTION_TYPE_ID " & Chr(129) & document.all.FDTID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME" & Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SCORE" & Chr(129) & document.all.TxtScore.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "RULE_ID"& Chr(129) & document.all.RULE_ID.innerText & Chr(129) & "1" & Chr(128)

		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()
		bRet = true
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If
	
	ExeSave = bRet
	window.returnValue = true
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
</script>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>" IsThisRequired="<%=isRequired%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Fraud Detection Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="FraudDetRuleSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="FDTID" value="<%=Request.QueryString("FDTID")%>">
<input type="hidden" NAME="FDRID" value="<%=Request.QueryString("FDRID")%>">

<%	

If FDRID <> "" Then
	If FDRID <> "NEW" Then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		SQLST = "SELECT FDR.*, R.RULE_TEXT " 
		SQLST=SQLST & "FROM FRAUD_DETECTION_RULE FDR, RULES R "
		SQLST=SQLST & "WHERE FDR.RULE_ID = R.RULE_ID(+) " 
		SQLST=SQLST & "AND FDR.FRAUD_DETECTION_RULE_ID = " & FDRID 
		Set oRS = oConn.Execute(SQLST)
	    If Not oRS.EOF Then
			RS_NAME = oRS("NAME")
			RS_RULE_ID = oRS("RULE_ID")
			RS_SCORE = oRS("SCORE")
			RS_RULE_TEXT= ReplaceQuotesInText(oRS("RULE_TEXT"))
		End If
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	End If
End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=s_DisplayMsg%></span>
</td>
</tr>
</table>

<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
	<td COLSPAN="3">Fraud Detection Rule ID:&nbsp;<span id="spanFDRID"><%=Request.QueryString("FDRID")%></span></td></tr>
<tr>
	<td width="75">Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="50" size="40" TYPE="TEXT" NAME="TxtName" VALUE="<%=RS_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td width="85">Score:<br><input ScrnInput="TRUE" CLASS="LABEL" SIZE="8" MAXLENGTH="8" TYPE="TEXT" NAME="TxtScore" VALUE="<%=RS_SCORE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>
<table class="Label" ONDRAGSTART="return false;">
<tr>
	<td width="40">
	<img NAME="BtnAttachRule" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
	<img NAME="BtnDetachRule" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
	</td>
	<td nowrap WIDTH="300">Rule Text:&nbsp;<span ID="RULE_TEXT" CLASS="LABEL" TITLE="<%=RS_RULE_TEXT%>"><%=TruncateText(RS_RULE_TEXT,RuleTextLen)%></span></td>
	<td>Rule ID:&nbsp;<span ID="RULE_ID"><%=RS_RULE_ID%></span></td>
</tr>	
<tr>
</tr>
</table>
</form>
</body>
</html>


