<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\CheckSharedRule.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	Dim SharedCount, SharedCountText, RID
	SharedCount = 0
	SharedCountText = "Ready"

	RID = Request.QueryString("RID")
		
	If RID <> "" Then
		If RID = "NEW" Then 
			SharedCount = 0
		Else
			SharedCount = CheckSharedRule(CLng(RID),true,true,1,false,false,0)
		End If
	End If	
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Rule</title>

<SCRIPT LANGUAGE=javascript SRC='..\lib\RefArray.js'></SCRIPT>	<!-- Reference arrays used by ETree -->
<SCRIPT LANGUAGE=javascript SRC='..\lib\ETree.js'></SCRIPT>		<!-- Expression tree evaluator class -->
<SCRIPT LANGUAGE=javascript>
<!--
function ParseExpr(strExpr)
{
	var retVal = true;
	var expr = ETreeFrom (strExpr);
	if (typeof expr != "object")
		retVal = false;
		
	return(retVal);
}

//-->
</SCRIPT>

<script LANGUAGE="VBScript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false


sub window_onload
<%	
	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScrnInputsReadOnly true,"DISABLED"
	SetScrnBtnsReadOnly true
<%	else
		if RID <> "" then
			if SharedCount <= 1 then %>
				document.all.ChkEdit.checked = true
				ChkEdit_OnClick
			<%else %>
				document.all.ChkEdit.checked = false
				ChkEdit_OnClick
				SetStatusInfoAvailableFlag(true)	
			<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
				  If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
						document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 color='Maroon'>+</Font>"
				<%	Else %>
						document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>"
				<%	End If
			end if						
		end if	'RID <> ""
	end if %>
End Sub

Sub PostTo(strURL)
	FrmRuleEditor.action = "RuleSearch-f.asp"
	FrmRuleEditor.method = "GET"
	FrmRuleEditor.target = "_parent"	
	FrmRuleEditor.submit
End Sub

sub SetRID(inRID)
	document.all.RID.value = inRID
	document.all.spanRuleID.innerText = inRID
end sub

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

function GetRID
	if document.all.RID.value <> "NEW" then
		GetRID = document.all.RID.value
	else
		GetRID = ""
	end if 
end function

function GetRIDText
	GetRIDText = document.all.TxtRuleText.value
end function

function GetRIDType
	GetRIDType = document.all.TxtRuleType.value
end function

function GetRIDComments
	GetRIDComments = document.all.TxtCommentsText.value
end function


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
dim nMaxRuleSize

strError = ""
<%
dim nMaxRuleSize
nMaxRuleSize = 2000
if instr(UCase(getInstanceName), "ESU") <> 0 then
	'	handle Esurance exception
	nMaxRuleSize = 4000
end if
%>
nMaxRuleSize = <%=nMaxRuleSize%>
If document.all.TxtRuleText.value = "" Then	strError = "Rule text is a required field." & VBCRLF
If Len(document.all.TxtRuleText.value) > nMaxRuleSize Then strError = strError & "Rule text cannot exceed " & CStr(nMaxRuleSize) & " characters."  & VBCRLF
If Len(document.all.TxtCommentsText.value) > 2000 Then strError = strError & "Comments text cannot exceed 2000 characters."  & VBCRLF
If strError <> "" Then
	ValidateScreenData = false
	MsgBox strError, 0 , "FNSNetDesigner"
Else
	ValidateScreenData = true
End If
End Function

Function ExeSave
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = false
		Exit Function
	End If
	
	If document.all.RID.value = "" Then
		ExeSave = false
		Exit Function
	End If
	
'	If CStr(document.body.getAttribute ("ScreenDirty")) = "YES" Then
		If ValidateScreenData = false Then 
			ExeSave = false
			Exit Function
		End If
		
		If (UCASE(document.all.TxtRuleType.value) = "ROUTING") Then
			FrmRuleEditor.submit
		ElseIf( ParseExpr(document.all.TxtRuleText.value) = false) then
			FrmRuleEditor.submit
	'		alert ("Error compiling SAVE expression '" + document.all.Expression.value + "'\n" + expr);		
		Else
			document.all.ValidClient.value = true
			FrmRuleEditor.submit		
		End If



		ExeSave = true
'	Else
'		document.all.SpanStatus.innerHTML = "Nothing to Save" 		
'		ExeSave = false
'	End If
End Function

Function ExeCopy
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.RID.value = "" Then
		ExeCopy = false
		Exit Function
	End If

	document.body.setAttribute "ScreenDirty","YES"
	document.all.RID.value = "NEW"
	document.all.SpanSharedCount.innerText = 0
	ExeCopy = ExeSave
End Function


sub SetScrnInputsReadOnly(bReadOnly, strNewClass)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnInput") = "TRUE" then
			document.all(iCount).readOnly = bReadOnly
			document.all(iCount).className = strNewClass
		end if
	next
end sub

sub SetScrnBtnsReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next
end sub

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"

	if document.all.ChkEdit.checked = true then
		SetScrnInputsReadOnly false,"LABEL"
		SetScrnBtnsReadOnly false
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScrnInputsReadOnly true,"DISABLED"
		SetScrnBtnsReadOnly true
		document.body.setAttribute "ScreenMode", "RO"				
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"	
end sub

sub Control_OnChange
<%	
	if CStr(Request.QueryString("MODE")) <> "RO" then 
%>
	document.body.setAttribute "ScreenDirty", "YES"	
<%	
end if 
%>
end sub

Sub BtnValidate_onclick()
		dim URLText
		'URLText=escape(document.all.TxtRuleText.value)
		URLText=encodeURIComponent(document.all.TxtRuleText.value)

		If URLText = "" Then
			MsgBox "Rule is not defined!"
			exit sub
		end if
		lret = window.showModalDialog ("validate.asp?RuleText=" & URLText,  "dialogWidth=1100px; dialogHeight=400px; center=yes")
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
		
End Sub

Sub RefCountRpt_onclick()
	If document.all.SpanSharedCount.innerText > 0 Then
		If document.all.RID.value <> "" And document.all.RID.value <> "NEW" Then
			paramID = document.all.RID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedRule=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
	Else
		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
	End If	
End Sub
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
<form name="FrmRuleEditor" method="POST" action="../rules/RuleEditorExe.asp" target="hiddenPage">
<% 'need to maintain these values in order to post back to the search tab %>

<input type="hidden" name="SearchRuleId" value="<%=Request.QueryString("SearchRuleId")%>">
<input type="hidden" name="SearchRuleText" value="<%=Request.QueryString("SearchRuleText")%>" ID="Hidden1">
<input type="hidden" name="SearchRuleType" value="<%=Request.QueryString("SearchRuleType")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" name="SearchComments" value="<%=Request.QueryString("SearchComments")%>" ID="Hidden2">
<input type="hidden" name="SearchUser" value="<%=Request.QueryString("SearchUser")%>" ID="Hidden3">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="RID" value="<%=Request.QueryString("RID")%>">
<input type="hidden" name="ValidClient" value=false>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Rule Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>


<%	dim RuleText, RuleType, RuleCommentsText
	RuleText = ""
	RuleType = ""
	RuleCommentsText = ""
	
	If RID <> "" then
		If RID <> "NEW" Then
			strExecute = "SELECT * FROM RULES WHERE RULE_ID = " & RID
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			' disable interface here...
			Set rs = Conn.Execute(strExecute)
			' enable
			
			if not rs.EOF then
				RuleType = rs("TYPE")		
				RuleText = rs("RULE_TEXT")
				RuleCommentsText = rs("COMMENTS")
			end if
			
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing		
		End If 'RID <> "NEW"

%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
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

<table CLASS="LABEL" ALIGN="CENTER" width="100%">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
<td>Rule ID:&nbsp;<span id="SpanRuleID" class="LABEL"><%=Request.QueryString("RID")%></span></td>
<td>Rule Type:
	<select ScrnBtn="TRUE" NAME="TxtRuleType" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetValidValuesHTML("RULES_TYPE",RuleType,true)%></select></td></tr>
 <td COLSPAN="2">Rule Text:<br>
 <textarea ScrnInput="TRUE" class="LABEL" MAXLENGTH="<%=nMaxRuleSize%>" name="TxtRuleText" cols="82" rows="15" style="overflow:hidden" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange"><%=RuleText%></textarea></td>
</tr>


<tr><td><button CLASS="StdButton" NAME="BtnValidate" ACCESSKEY="V"  ID="ButtonVal"><u>V</u>alidate</button></td>


</table>


<table CLASS="LABEL" ALIGN="CENTER" width="100%" ID="Table2">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
 <td COLSPAN="2">Comments:<br>
 <textarea ScrnInput="TRUE" class="LABEL" MAXLENGTH="2000" name="TxtCommentsText" cols="82" rows="10" style="overflow:hidden" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange"><%=RuleCommentsText%></textarea></td>
</tr>
</table>


<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No rule selected.
</div>


<% End If %>


</body>
</form>
</html>

