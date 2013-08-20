<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	Response.Expires=0 %>
<!--#include file="..\lib\ZIP.inc"-->
<html>
<head>
<title>iNetPolicy Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
var g_StatusInfoAvailable = false;
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<% if Request.QueryString("MODE") = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<% end if
	%>
	document.all.txtAction.Value = "<%=Request.QueryString("txtAction")%>"
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "iNetPolicySearch-f.asp"
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
dim cErrMsg

cErrMsg = ""
if document.all.txtClientCode.value = "" then
	cErrMsg = "Client Code is a required field." & VbCrlf
end if
if document.all.txtPolicyID.value = "" then
	cErrMsg = cErrMsg & "Policy ID is a required field." & VbCrlf
end if
If cErrMsg = "" Then
	ValidateScreenData = true		
Else
	msgbox cErrMsg, 0, "FNSDesigner"
	ValidateScreenData = false
End If
End Function

Function ExeSave()
dim cResult

cResult = ""
if not ValidateScreenData then 
	ExeSave = false
	exit function
end if
if document.all.txtAction.Value = "UPDATE" then
cResult = "UPDATE INETPOLICY SET "
cResult = cResult & "CARRIER_NAME='" & document.all.txtCarrierName.value & "',"
cResult = cResult & "INSURED_NAME='" & document.all.txtInsuredName.value & "',"
cResult = cResult & "ADDRESS_LINE1='" & document.all.txtAddress1.value & "',"
cResult = cResult & "ADDRESS_LINE2='" & document.all.txtAddress2.value & "',"
cResult = cResult & "ADDRESS_CITY='" & document.all.City.value & "',"
cResult = cResult & "ADDRESS_STATE='" & document.all.State.value & "',"
cResult = cResult & "ADDRESS_ZIP='" & document.all.Zip.value & "' WHERE "
cResult = cResult & "CLIENT_CD='" & document.all.txtClientCode.value & "' AND "
cResult = cResult & "POLICY_IDENTIFIER='" & document.all.txtPolicyID.value & "'"
else
cResult = "INSERT INTO INETPOLICY VALUES ("
cResult = cResult & "'" & document.all.txtClientCode.value & "','"
cResult = cResult & document.all.txtPolicyID.value & "','"
cResult = cResult & document.all.txtCarrierName.value & "','"
cResult = cResult & document.all.txtInsuredName.value & "','"
cResult = cResult & document.all.txtAddress1.value & "','"
cResult = cResult & document.all.txtAddress2.value & "','"
cResult = cResult & document.all.City.value & "','"
cResult = cResult & document.all.State.value & "','"
cResult = cResult & document.all.Zip.value & "')"
end if
document.all.txtSaveData.Value = cResult
document.all.FrmDetails.Submit()
ExeSave = true
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

<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» iNetPolicy Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="iNetPolicySave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="txtSaveData">
<input TYPE="HIDDEN" NAME="txtAction">

<input type="hidden" name="ClientCode" value="<%=Request.QueryString("ClientCode")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="PolicyID" value="<%=Request.QueryString("PolicyID")%>">
<input type="hidden" name="CarrierName" value="<%=Request.QueryString("CarrierName")%>">
<input type="hidden" name="InsuredName" value="<%=Request.QueryString("InsuredName")%>">
<%	
Dim cClientCode, cAction, oConn, cSQL, oRS, nRec
dim aKeys

cAction = Request.QueryString("txtAction")
nRec = Request.QueryString("recBM")
if cAction <> "NEW" and nRec <> "" then
	aKeys = Split(nRec, "^")
	cClientCode = aKeys(0)
	cPolicyID = aKeys(1)
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "SELECT * "
	cSQL = cSQL & "FROM INETPOLICY WHERE CLIENT_CD='" & cClientCode & "' AND POLICY_IDENTIFIER='" & cPolicyID & "'"
	Set oRS = oConn.Execute(cSQL)
	If Not oRS.EOF then
		cCarrierName = oRS("CARRIER_NAME")
		cInsuredName = oRS("INSURED_NAME")
		cAddress1 = oRS("ADDRESS_LINE1")
		cAddress2 = oRS("ADDRESS_LINE2")
		cCity = oRS("ADDRESS_CITY")
		cState = oRS("ADDRESS_STATE")
		cZip = oRS("ADDRESS_ZIP")
	end if	
	oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
end if	
if nRec <> "" then	
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
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr><td>
<table class="LABEL">
<tr>
<td CLASS="LABEL">Client Code:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="3" SIZE="3" TYPE="TEXT" NAME="txtClientCode" VALUE="<%=cClientCode%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<tr>
<td CLASS="LABEL">Policy Id:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" SIZE="80" TYPE="TEXT" NAME="txtPolicyID" VALUE="<%=cPolicyID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Carrier Name:<br><input ScrnInput="TRUE" size="80" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="txtCarrierName" VALUE="<%=cCarrierName%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
<td CLASS="LABEL">Insured Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" size="80" TYPE="TEXT" NAME="txtInsuredName" VALUE="<%=cInsuredName%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Address 1:<br><input type="text" Size="80" MAXLENGTH="80" Scrninput="TRUE" NAME="txtAddress1" VALUE="<%=cAddress1%>" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Address 2:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" size="80" TYPE="TEXT" NAME="txtAddress2" VALUE="<%=cAddress2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Zip Code:<br><input ScrnInput="TRUE" CLASS="LABEL" size="9" MAXLENGTH="9" TYPE="TEXT" NAME="Zip" VALUE="<%=cZip%>"></td>
	<td CLASS="LABEL">City:<br><input size="40" CLASS="READONLY" READONLY MAXLENGTH="40" TYPE="TEXT" NAME="City" VALUE="<%=cCity%>" ></td>
	<td CLASS="LABEL">State:<br><input CLASS="READONLY" size="2" READONLY MAXLENGTH="2" TYPE="TEXT" NAME="State" VALUE="<%=cState%>" ></td>
	</tr>
	</table>
	
 <% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No record selected.
</div>
<% End If %>
</form>
</body>
</html>


