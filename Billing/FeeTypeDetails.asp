<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%Response.Expires=0 
	Dim SharedCount, SharedCountText, FID
	SharedCount = 0
	SharedCountText = "Ready"
	
	FID	= CStr(Request.QueryString("FID"))

	If FID <> "" Then
		If FID = "NEW" Then 
			SharedCount = 0
		'Else
		'	SharedCount = CheckSharedAttribute(CLng(FID),True,True,1,False,False,0)
		End If
	End If	
	
If FID <> "" Then
RSACCNT_HRCY_STEP_ID = Request.QueryString("AHSID")
	If FID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM FEE_TYPE WHERE FEE_TYPE_ID = " & FID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSFEE_TYPE_ID = RS("FEE_TYPE_ID")
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
<title>Fee Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if FID <> "" then %>
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
	FrmDetails.action = "FeeTypeSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateFID(inFID)
	document.all.FID.value = inFID
	document.all.spanFID.innerText = inFID
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

Function GetFID
	if document.all.FID.value <> "NEW" then
		GetFID = document.all.FID.value
	else
		GetFID = ""
	end if 
End Function

Function GetFIDName
	GetFIDName = document.all.TxtName.value
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
	If document.all.TxtName.value = "" then
		msgbox "Accnt Hrcy Step ID is a required field.", 0, "FNSDesigner"
		ValidateScreenData = false
		exit function
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
	
	if document.all.FID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.FID.value = "NEW"
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
	
	if document.all.FID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.FID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "FEE_TYPE_ID"& Chr(129) & document.all.FID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtNAME.value & Chr(129) & "1" & Chr(128)
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
'	If document.all.SpanSharedCount.innerText > 0 Then
'		If document.all.FID.value <> "" And document.all.FID.value <> "NEW" Then
'			paramID = document.all.FID.value
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
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Fee Type Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<form Name="FrmDetails" METHOD="POST" ACTION="FeeTypeSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" name="SearchFID" value="<%=Request.QueryString("SearchFID")%>">
<input type="hidden" name="SearchNAME" value="<%=Request.QueryString("SearchNAME")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="FID" value="<%=Request.QueryString("FID")%>" >

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
If FID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<!--
<td WIDTH="14">
<img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count">
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">
:<span id="SpanSharedCount"><%=SharedCount%></span>
</td>
-->
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</td>
<td>
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING=0 CELLSPACING=0 >
<tr>
<td>
<table class="LABEL">
	<tr>
	<td COLSPAN=5 CLASS=LABEL>Fee Type ID:&nbsp<span id="spanFID"><%=Request.QueryString("FID")%></span></td>
	</tr>
	<tr>
	<td CLASS=LABEL>Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=255 size="80" TYPE="TEXT" NAME="TxtNAME" VALUE="<%=RSNAME%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</TR>
	</TD>
	</TABLE>
</TABLE>
<% Else %>
<DIV style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No fee type selected.
</DIV>
<% End If %>
</form>
</body>
</html>


