<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	
Response.Expires = 0 
Response.AddHeader  "Pragma", "no-cache"
Response.Buffer = true
RuleTextLen = 30
RSAHSID = Request.QueryString("AHSID")
	
Dim FDTID

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
FDTID = Request.QueryString("FDTID")
If FDTID <> "" Then
	If FDTID <> "NEW" Then
	    SQLST = "SELECT * FROM " 
		SQLST =SQLST & "FRAUD_DETECTION_TYPE FDT, ACCOUNT_HIERARCHY_STEP AHS " 
		SQLST =SQLST & "WHERE FDT.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID"
		SQLST =SQLST & " AND FRAUD_DETECTION_TYPE_ID =" & FDTID 
				 
		Set oRS = oConn.Execute(SQLST)
		If Not oRS.EOF Then
		    RSDESCRIPTION = ReplaceQuotesInText(oRS("DESCRIPTION"))
			RSLOB = oRS("LOB_CD")			
			RSTHRESHOLD = oRS("THRESHOLD")
		End If
		oRS.Close
	End If
%>
	
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Fraud Detection Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="javascript">

var g_StatusInfoAvailable = false;

function CFraudDetRuleSearchObj()
{
	this.Selected = false;
}

var FraudDetRuleSearchObj = new CFraudDetRuleSearchObj();
</script>

<script LANGUAGE="JScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 200;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 180;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
<%
If FDTID <> "" Then
%>		
	document.all.LOB_CD.value = "<%= RSLOB %>"
<%
end if
%>	
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub UpdateSpanText(SPANID,inText)
	If Len(inText) < <%=RuleTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid (inText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Sub UpdateFDTID(inFDTID)
	document.all.FDTID.value = inFDTID
	document.all.spanFDTID.innerText = inFDTID
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

Function GetFDTID
	if document.all.FDTID.value <> "NEW" then
		GetFDTID = document.all.FDTID.value
	else
		GetFDTID = ""
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
	IF CStr(document.all.getAttribute("IsThisRequired")) = "Y" Then
		f_CheckIsThisRequired = true
	ELSE
		f_CheckIsThisRequired = False
	END IF
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	If  document.all.TxtDescription.value = "" then
		MsgBox "Description is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	If  document.all.LOB_CD.value = "" then
		MsgBox "LOB is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	If document.all.TxtThreshold.value = "" then
		MsgBox "Threshold is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	ValidateScreenData = true
End Function

Function GetSelectedFDRID
	GetSelectedFDRID = document.frames("DataFrame").GetSelectedFDRID
End Function

Sub ExeNewBranchRule

	dim FDTID, FDRID, MODE
	
	If Not InEditMode Then
		Exit Sub
	End If
	If document.all.FDTID.value = "" Or document.all.FDTID.value = "NEW" Then
		Exit Sub
	End If

	FDRID = "NEW"
	FDTID = document.all.FDTID.value
	MODE = document.body.getAttribute("ScreenMode")

	FraudDetRuleSearchObj.Selected = false
	strURL = "FraudDetRuleModal.asp?FDTID=" & FDTID & "&FDRID=" & FDRID & "&MODE=" & MODE 	
	showModalDialog strURL, FraudDetRuleSearchObj, "center:yes;status:no;help:no" 
	If FraudDetRuleSearchObj.Selected Then 
		Refresh
	end if
End Sub

Sub Refresh
	FDTID = document.all.FDTID.value
	document.all.tags("IFRAME").item("DataFrame").src = "FraudDetDetailsData.asp?FDTID=" & FDTID
End Sub

Sub ExeEditBranchRule
	dim FDRID, FDTID

	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.FDTID.value = "" Or document.all.FDTID.value = "NEW" Then
		Exit Sub
	End If

	FDRID = GetSelectedFDRID
	FDTID = document.all.FDTID.value
	
	If FDRID <> "" Then
		FraudDetRuleSearchObj.Selected = false
		strURL = "FraudDetRuleModal.asp?FDTID=" & FDTID & "&FDRID=" & FDRID & "&MODE=" & MODE 	
		showModalDialog  strURL, FraudDetRuleSearchObj, "center"
		If FraudDetRuleSearchObj.Selected Then
			Refresh
		end if
	Else
		MsgBox "Please select a Fraud Detection Rule to Edit.", 0, "FNSNet Designer"		
	End If
	
End Sub

Sub ExeRemoveBranchRule
	dim FDRID, sResult

	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.FDTID.value = "" Or document.all.FDTID.value = "NEW" Then
		Exit Sub
	End If

	FDRID = GetSelectedFDRID
	FDTID = document.all.FDTID.value
	
	If FDRID <> "" Then
		sResult = sResult & FDRID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmDetails.action = "FraudDetRuleSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a Fraud Detection Rule to Remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub

Function InEditMode
	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		InEditMode = false
	End If
End Function

Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.FDTID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	if ValidateScreenData = false then 
		ExeSave = false
		exit function
	end if

	If document.all.FDTID.value = "NEW" then
		document.all.TxtAction.value = "INSERT"
	else
		document.all.TxtAction.value = "UPDATE"
	end if
	sResult = "FRAUD_DETECTION_TYPE_ID"& Chr(129) & document.all.FDTID.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & "<%=RSAHSID%>" & Chr(129) & "1" & Chr(128)
	sResult = sResult & "LOB_CD" & Chr(129) & document.all.LOB_CD.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "THRESHOLD" & Chr(129) & document.all.TxtThreshold.value & Chr(129) & "1" & Chr(128)
	
	document.all.TxtSaveData.Value = sResult
	FrmDetails.action = "FraudDetSave.asp"
	FrmDetails.method = "POST"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit
	bRet = true
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
<!--#include file="..\lib\Help.asp"

Sub window_onunload

End Sub

-->
</script>
<!--#include file="..\lib\BABtnControl.inc"-->

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
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

<form Name="FrmDetails" METHOD="POST" ACTION="FraudDetSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="FDTID" value="<%=Request.QueryString("FDTID")%>">
<input type="hidden" NAME="AHSID" value="<%=RSAHSID%>" ID="Hidden1">

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td colspan="2">Fraud Detection Type ID:&nbsp;<span id="spanFDTID"><%=Request.QueryString("FDTID")%></span></td></tr>
<tr>
	<td>Description:<br><input ScrnInput="TRUE" MAXLENGTH="128" CLASS="LABEL" size="65" TYPE="TEXT" NAME="TxtDescription" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>LOB:<br>
	<select STYLE="WIDTH:150" NAME="LOB_CD" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	<%
	cSQL = "SELECT * FROM LOB"
	Set oRS2 = oConn.Execute(cSQL)
	Do WHile Not oRS2.EOF
	%>
		<option VALUE="<%= oRS2("LOB_CD") %>"><%= oRS2("LOB_NAME") %>
	<%
		oRS2.MoveNext
	Loop
	oRS2.Close
	%>
	</select>
	</td>
</tr>
<tr>
	<td>Threshold:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="10" TYPE="TEXT" NAME="TxtThreshold" VALUE="<%=RSTHRESHOLD%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text1"></td>
</tr>
</table>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Fraud Detection Rules</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<span class="Label" ID="SPANDATA">Retrieving...</span>
<fieldset id="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&amp;HIDEATTACH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="BABtnControl" type="text/x-scriptlet"></object>
<iframe width="100%" height="0" name="DataFrame" src="FraudDetDetailsData.asp?<%=Request.QueryString%>">
</fieldset>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Fraud Detection record selected.
</div>

<% End If 
Set oRS = Nothing
oConn.Close
Set oConn = Nothing
%>
</form>
</body>
</html>


