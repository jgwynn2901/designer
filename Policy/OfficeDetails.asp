<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\CheckSharedOffice.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	Dim SharedCount, SharedCountText, OID
	SharedCount = 0
	SharedCountText = "Ready"

	OID = Request.QueryString("OID")
	
	If OID <> "" Then
		If OID = "NEW" Then 
			SharedCount = 0
		Else
			SharedCount = CheckSharedOffice(CLng(OID),true,true,1,false,false,0)
		End If
	End If	

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Office</title>

<script>
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
<script LANGUAGE="VBScript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false


sub window_onload
<%	
	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScrnInputsReadOnly true,"DISABLED"
	SetScrnBtnsReadOnly true
<%	else
		if OID <> "" then
			if SharedCount <= 1 then %>
	document.all.ChkEdit.checked = true
	ChkEdit_OnClick

<%			else %>
	document.all.ChkEdit.checked = false
	ChkEdit_OnClick
	SetStatusInfoAvailableFlag(true)	
<%
				SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
						
		end if	'OID <> ""

	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "OfficeSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
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

Sub UpdateOID(inOID)
	document.all.OID.value = inOID
	document.all.spanOID.innerText = inOID
End Sub


function GetOID
	if document.all.OID.value <> "NEW" then
		GetOID = document.all.OID.value
	else
		GetOID = ""
	end if 
end function

function GetOIDNumber
	GetOIDNumber = document.all.TxtNumber.value
end function

function GetOIDState
	GetOIDState = document.all.TxtState.value
end function

function GetOIDType
	GetOIDType = document.all.TxtType.value
end function

function GetOIDZip
	GetOIDZip = document.all.TxtZip.value
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
errStr = ""
	If document.all.TxtState.value = "" Then errStr =  "State is a required field." & VBCRLF
	If document.all.TxtType.value = "" Then errStr = errStr & "Type is a required field." & VBCRLF
	If document.all.TxtZip.value = "" Then errStr = errStr & "Zip is a required field." & VBCRLF


	If document.all.TxtNumber.value = "" Then 
		errStr = errStr & "Number is a required field." & VBCRLF
	elseif IsNumeric(document.all.TxtNumber.value) = false then
		errStr = errStr & "Number must be numeric." & VBCRLF
	end If

	If document.all.TxtLAT.value <> "" And IsNumeric(document.all.TxtLAT.value) = false then
			errStr =  errStr & "LAT must be numeric." & VBCRLF
	End If
	If document.all.TxtLON.value <> "" And IsNumeric(document.all.TxtLON.value) = false then
			errStr =  errStr & "LON must be numeric." & VBCRLF
	End If

	If errstr = "" Then
		ValidateScreenData = true
	Else
		MsgBox errstr, 0 , "FNSNetDesigner"
		ValidateScreenData = false
	End If	

End Function

Function ExeCopy
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.OID.value = "" Then
		ExeCopy = false
		Exit Function
	End If

	document.body.setAttribute "ScreenDirty","YES"
	document.all.OID.value = "NEW"
	document.all.SpanSharedCount.innerText = 0
	ExeCopy = ExeSave
End Function


Function ExeSave
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = false
		Exit Function
	End If

	If document.all.OID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.OID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "PK_OFFICE"& Chr(129) & document.all.OID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OFFICE_NUMBER"& Chr(129) & document.all.TxtNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATUS"& Chr(129) & document.all.TxtStatus.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OFFICE_TYPE"& Chr(129) & document.all.TxtType.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OFFICE_NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_1"& Chr(129) & document.all.TxtAddress1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_2"& Chr(129) & document.all.TxtAddress2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.TxtCity.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.TxtState.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.TxtZip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX"& Chr(129) & document.all.TxtFax.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CONTACT_F_NAME"& Chr(129) & document.all.TxtFirstName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CONTACT_L_NAME"& Chr(129) & document.all.TxtLastName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CONTACT_TITLE"& Chr(129) & document.all.TxtTitle.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LAT"& Chr(129) & document.all.TxtLAT.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LON"& Chr(129) & document.all.TxtLON.value & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "OfficeDetailsSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
			
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If
	
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
		If document.all.OID.value <> "" And document.all.OID.value <> "NEW" Then
			paramID = document.all.OID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedOffice=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
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
<body topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<form name="FrmDetails" method="GET" action="OfficeDetailsSave.asp" target="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchOID" value="<%=Request.QueryString("SearchOID")%>">
<input type="hidden" name="SearchNumber" value="<%=Request.QueryString("SearchNumber")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchOType" value="<%=Request.QueryString("SearchOType")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="OID" value="<%=Request.QueryString("OID")%>">

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Office Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>


<%	If OID <> "" then
		If OID <> "NEW" Then
			strExecute = "SELECT * FROM Office WHERE PK_OFFICE = " & OID
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			Set rs = Conn.Execute(strExecute)

			if not rs.EOF then
				RSNUMBER = rs("OFFICE_NUMBER")		
				RSSTATUS = rs("STATUS")		
				RSTYPE = ReplaceQuotesInText(rs("OFFICE_TYPE"))
				RSNAME = ReplaceQuotesInText(rs("OFFICE_NAME"))
				RSADDRESS1 = ReplaceQuotesInText(rs("ADDRESS_1"))
				RSADDRESS2 = ReplaceQuotesInText(rs("ADDRESS_2"))
				RSCITY = rs("CITY")
				RSSTATE = rs("STATE")
				RSZIP = rs("ZIP")
				RSPHONE = rs("PHONE")
				RSFAX = rs("FAX")
				RSFIRSTNAME = ReplaceQuotesInText(rs("CONTACT_F_NAME"))
				RSLASTNAME = ReplaceQuotesInText(rs("CONTACT_L_NAME"))
				RSTITLE = ReplaceQuotesInText(rs("CONTACT_TITLE"))
				RSLAT = rs("LAT")
				RSLON = rs("LON")
			end if
			
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing		
		End If 'OID <> "NEW"

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

<table CLASS="LABEL" ALIGN="LEFT">
<tr></tr> 
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td>Office ID:&nbsp;<span id="SpanOID" class="LABEL"><%=Request.QueryString("OID")%></span></td></tr>
<tr><td>Number:<br><input type="text" ScrnInput="TRUE" size=12 class="LABEL" name="TxtNumber" maxlength=10  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSNUMBER%>"></input></td>
	<td>Status:<br><input type="text" ScrnInput="TRUE" size=2 class="LABEL" name="TxtStatus" maxlength=1  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSSTATUS%>"></input></td>
    <td>Type:<br><input type="text" ScrnInput="TRUE"  size=6 class="LABEL" name="TxtType" maxlength=5  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSTYPE%>"></input></td>
	<td COLSPAN=2>Name:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtName" maxlength=45  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSNAME%>"></input></td>    
</tr>
<tr>
	<td COLSPAN=3>Address 1:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtAddress1" maxlength=80  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSADDRESS1%>"></input></td>
    <td COLSPAN=2>Address 2:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtAddress2" maxlength=80  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSADDRESS2%>"></input></td>
</tr>
<tr><td COLSPAN=3>City:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtCity" maxlength=40  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSCITY%>"></input></td>
    <td>State:<br><SELECT ScrnBtn="TRUE" NAME=TxtState CLASS=LABEL ONCHANGE="VBScript::Control_OnChange"><OPTION VALUE=""><!--#include file="..\lib\states.asp"--></SELECT></td>
    <td>Zip:<br><input type="text" ScrnInput="TRUE" size=10 class="LABEL" name="TxtZip" maxlength=9  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSZIP%>"></input></td>
</tr>
<tr><td COLSPAN=3>Phone:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtPhone" maxlength=14  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSPHONE%>"></input></td>
    <td COLSPAN=2>Fax:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtFax" maxlength=10  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSFAX%>"></input></td>
</tr>
<tr><td COLSPAN=3>Contact First Name:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtFirstName" maxlength=40  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSFIRSTNAME%>"></input></td>
    <td COLSPAN=2>Contact Last Name:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtLastName" maxlength=80  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSLASTNAME%>"></input></td>
</tr>
<tr><td COLSPAN=3>Contact Title:<br><input type="text" ScrnInput="TRUE" size=32 class="LABEL" name="TxtTitle" maxlength=80  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSTITLE%>"></input></td>
</tr>

<tr><td>LAT:<br><input type="text" ScrnInput="TRUE" size=14 class="LABEL" name="TxtLAT" maxlength=10  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSLAT%>"></input></td>
    <td COLSPAN=2>LON:<br><input type="text" ScrnInput="TRUE" size=14 class="LABEL" name="TxtLON" maxlength=10  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSLON%>"></input></td>
</tr>
</table>

<%		If Not IsNull(RSSTATE) Then
			If  CStr(RSSTATE) <> "" Then %>
<SCRIPT LANGUAGE="VBScript">
	SelectOption document.all.TxtState,"<%=CStr(RSSTATE)%>"
</SCRIPT>
<%			End If
		End If  %>

<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No office selected.
</div>


<% End If %>


</body>
</form>
</html>

