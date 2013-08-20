<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\CheckSharedRoutingAddress.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	Dim SharedCount, SharedCountText, RAID
	SharedCount = 0
	SharedCountText = "Ready"

	RAID = Request.QueryString("RAID")
	
	If RAID <> "" Then
		If RAID = "NEW" Then 
			SharedCount = 0
		Else
			SharedCount = CheckSharedRoutingAddress(CLng(RAID),true,true,1,false,false,0)
		End If
	End If	

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Routing Address</title>

<script LANGUAGE="VBScript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false


sub window_onload
<%	
	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScrnInputsReadOnly true,"DISABLED"
	SetScrnBtnsReadOnly true
<%	else
		if RAID <> "" then
			if SharedCount <= 1 then %>
				document.all.ChkEdit.checked = true
				ChkEdit_OnClick

		<%	else %>
				document.all.ChkEdit.checked = false
				ChkEdit_OnClick
				SetStatusInfoAvailableFlag(true)	
			<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
				If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
					document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 Color='Maroon'>+</Font>"
				<%Else %>
					document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>"
				<%End If			
			end if					
		end if	'RAID <> ""
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "RoutingAddressSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

sub SetRAID(inRAID)
	document.all.RAID.value = inRAID
	document.all.spanRAID.innerText = inRAID
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

Sub UpdateRAID(inRAID)
	document.all.RAID.value = inRAID
	document.all.spanRAID.innerText = inRAID
End Sub


function GetRAID
	if document.all.RAID.value <> "NEW" then
		GetRAID = document.all.RAID.value
	else
		GetRAID = ""
	end if 
end function

function GetRAIDDescription
	GetRAIDDescription = document.all.TxtDescription.value
end function

function GetRAIDState
	GetRAIDState = document.all.TxtState.value
end function

function GetRAIDFIPS
	GetRAIDFIPS = document.all.TxtFIPS.value
end function

function GetRAIDZip
	GetRAIDZip = document.all.TxtZip.value
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
	If document.all.TxtState.value = "" Then
		MsgBox "State is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		Exit Function
	ElseIf  document.all.TxtFIPS.value = "" Then
		MsgBox "FIPS is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		Exit Function
	ElseIf  document.all.TxtZip.value = "" Then
		MsgBox "Zip is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		Exit Function
	End If
	ValidateScreenData = true
End Function

Function ExeCopy
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.RAID.value = "" Then
		ExeCopy = false
		Exit Function
	End If

	document.body.setAttribute "ScreenDirty","YES"
	document.all.RAID.value = "NEW"
	document.all.SpanSharedCount.innerText = 0
	ExeCopy = ExeSave
End Function


Function ExeSave
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = false
		Exit Function
	End If

	If document.all.RAID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.RAID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "ROUTINGADDRESS_ID"& Chr(129) & document.all.RAID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.TxtState.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FIPS"& Chr(129) & document.all.TxtFIPS.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.TxtZip.value & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "RoutingAddressDetailsSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
			
		bRet = true
		
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If
	
	ExeSave = bRet
	
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
<%	if CStr(Request.QueryString("MODE")) <> "RO" then %>
	document.body.setAttribute "ScreenDirty", "YES"	
<%	end if %>
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
		If document.all.RAID.value <> "" And document.all.RAID.value <> "NEW" Then
			paramID = document.all.RAID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedRoutingAddress=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
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
</script> 

</head>
<body topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<form name="FrmDetails" method="GET" action="RoutingAddressDetailsSave.asp" target="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchRAID" value="<%=Request.QueryString("SearchRAID")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchFIPS" value="<%=Request.QueryString("SearchFIPS")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="RAID" value="<%=Request.QueryString("RAID")%>">

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Routing Address Details</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>


<%	If RAID <> "" then
		If RAID <> "NEW" Then
			strExecute = "SELECT * FROM ROUTINGADDRESS WHERE ROUTINGADDRESS_ID = " & RAID
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			Set rs = Conn.Execute(strExecute)

			if not rs.EOF then
				RSDESCRIPTION = rs("DESCRIPTION")		
				RSSTATE = rs("STATE")
				RSFIPS = rs("FIPS")
				RSZIP = rs("ZIP")
			end if
			
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing		
		End If 'RAID <> "NEW"

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
<input ScrnBtn ="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>

<table CLASS="LABEL" ALIGN="CENTER" width="100%" >
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td>Routing Address ID:&nbsp;<span id="SpanRAID" class="LABEL"><%=Request.QueryString("RAID")%></span></td></tr>
<tr><td>Description:<br><input type="text" ScrnInput="TRUE" maxlength=255 size=100 class="LABEL" name="TxtDescription" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSDESCRIPTION%>"></input></td></tr>
<tr><td>State:<br><input type="text" ScrnInput="TRUE" maxlength=255 size=100 class="LABEL" name="TxtState" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSSTATE%>"></input></td></tr>
<tr><td>FIPS:<br><input type="text" ScrnInput="TRUE" maxlength=255 size=100 class="LABEL" name="TxtFIPS" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSFIPS%>"></input></td></tr>
<tr><td>Zip:<br><input type="text" ScrnInput="TRUE" maxlength=255 size=100 class="LABEL" name="TxtZip"  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSZIP%>"></input></td></tr>
</table>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No routing address selected.
</div>


<% End If %>


</body>
</form>
</html>

