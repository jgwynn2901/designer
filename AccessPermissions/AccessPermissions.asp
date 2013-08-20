<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<%	
	Response.Expires=0 
	Dim UID, GID
	
	UID =  CStr(Request.QueryString("UID"))
	GID =  CStr(Request.QueryString("GID"))
	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Access Permissions</title>
<link rel="stylesheet" type="text/css" href="../users/..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
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
	If  document.all.TxtFunction.value = "" then
		MsgBox "Function is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	If  document.all.TxtAccessType.value = "" then
		MsgBox "Access Type is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
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

	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		document.all.TxtAction.value = "INSERT"
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "ACCESS_ID"& Chr(129) & "" & Chr(129) & "1" & Chr(128)
		sResult = sResult & "USER_ID"& Chr(129) & document.all.UID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "GROUP_ID"& Chr(129) & document.all.GID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FUNCTION_ID"& Chr(129) & document.all.TxtFunction.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCESSTYPE_ID"& Chr(129) & document.all.TxtAccessType.value & Chr(129) & "1" & Chr(128)

		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmAccessPermissions.Submit()
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If
	
	ExeSave = bRet
	
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		UpdateStatus("Ready")
	end if
end sub

sub SetScreenFieldsReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
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
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Access Permissions for <%=Request.QueryString("TITLE")%></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmAccessPermissions" METHOD="POST" ACTION="AccessPermissionsSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="UID" value="<%=Request.QueryString("UID")%>">
<input type="hidden" NAME="GID" value="<%=Request.QueryString("GID")%>">

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="../users/..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
<td>Function:<br>
	<select ID="TxtFunction" CLASS="LABEL" ScrnBtn="TRUE" NAME="TxtFunction">
	<%
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "Select * From FUNCTION Order By FUNCTION_NM"
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
	<option VALUE="<%=oRS("FUNCTION_ID")%>"><%=oRS("FUNCTION_NM")%>
	<%
	oRS.MoveNext
	Loop
	oRS.CLose
	oConn.close
	set oConn = nothing
	set oRS = nothing
	%>
</td>	
<td>AccessType:<br><select ScrnBtn="TRUE" NAME="TxtAccessType" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("ACCESSTYPE","ACCESSTYPE_ID","ACCESSTYPE","",true)%></select></td>
</tr>
</table>
</form>
</body>
</html>


