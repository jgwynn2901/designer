<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	Response.Expires=0 
	Dim LUCID, LUTID
	LUCID =  CStr(Request.QueryString("LUCID"))
	LUTID =  CStr(Request.QueryString("LUTID"))
	'KROS-0039 (for Tower, Set maxlength of "Code" field to 255 and allow null in "Value" field)
	Dim IsTOWER, CodeMaxLength
	if left(getInstanceName,3) = "TOW" then
		IsTOWER = true
		CodeMaxLength = 255
	else
		IsTOWER = false
		CodeMaxLength = 40
	end if
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Lookup Code Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
End Sub

Sub UpdateLUCID(inLUCID)
	document.all.LUCID.value = inLUCID
	document.all.spanLUCID.innerText = inLUCID
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
	If  document.all.TxtCode.value = "" then
		MsgBox "Code is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
   'KROS-0039 (for Tower, Set maxlength of "Code" field to 255 and allow null in "Value" field)
	If Not <%=IsTOWER%> Then
		If  document.all.TxtValue.value = "" Then
			MsgBox "Value is a required field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		End If
	End If
	
	If document.all.TxtSequence.value <> "" Then
		If IsNumeric(document.all.TxtSequence.value) = false then
			MsgBox "Please enter a number in the Sequence field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		End If
	End If
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
	
	If document.all.LUTID.value = "" Then
		ExeSave = false
		exit function
	End If
		
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.LUCID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "LU_ID"& Chr(129) & document.all.LUCID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LU_TYPE_ID"& Chr(129) & document.all.LUTID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CODE"& Chr(129) & document.all.TxtCode.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "VALUE"& Chr(129) & document.all.TxtValue.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQUENCE"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)

		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()
		bRet = true
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If
	
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
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Lookup Code Details</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="LookupCodeSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="LUTID" value="<%=Request.QueryString("LUTID")%>" >
<input type="hidden" NAME="LUCID" value="<%=Request.QueryString("LUCID")%>" >

<%	
If LUCID <> "" Then
	If LUCID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM LU_CODE WHERE LU_ID = " & LUCID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSCODE = ReplaceQuotesInText(RS("CODE"))
			RSVALUE = ReplaceQuotesInText(RS("VALUE"))
			RSSEQUENCE= RS("SEQUENCE")
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>
<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
<td>Lookup Code ID:&nbsp<span id="spanLUCID"><%=Request.QueryString("LUCID")%></span></td>
<td>&nbsp</td>
</tr>
<tr>
<td>Code:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="<%=CodeMaxLength%>" SIZE=42 TYPE="TEXT" NAME="TxtCode" VALUE="<%=RSCODE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Sequence:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=11 TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td colspan=3>Value:<br><input ScrnInput="TRUE" CLASS="LABEL" SIZE=65 MAXLENGTH=255 TYPE="TEXT" NAME="TxtValue" VALUE="<%=RSVALUE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>
</form>
</body>
</html>


