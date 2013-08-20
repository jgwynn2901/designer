<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	
Response.Expires=0

dim cAccVendorID, oConn, oRS, cSQL
dim nVID, nNID, nSEQ, nAHSID, cLOB, nST, nCM
dim cError, lUpdateOK, cServiceType
dim lIsEdit

lIsEdit = len(Request.QueryString("EDIT")) <> 0
nVID = Request.QueryString("VID")
nNID = Request.QueryString("NID")
nSEQ = Request.QueryString("SEQ")	
nAHSID = Request.QueryString("AHSID")
cLOB = Request.QueryString("LOB")
nST = Request.QueryString("ST")
nCM = Request.QueryString("CM")

cAccVendorID = trim(Request.QueryString("AVID"))
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
cSQL = "SELECT * FROM SERVICE_TYPE WHERE SERVICE_TYPE_ID = " & nST
Set oRS = oConn.Execute(cSQL)
cServiceType = oRS("TYPE")
oRS.close
if lIsEdit then
	cSQL = "SELECT * FROM ACCOUNT_VENDOR WHERE ACCOUNT_VENDOR_ID = " & cAccVendorID
	Set oRS = oConn.Execute(cSQL)
end if	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Add or Edit Vendor/Network to Account</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
function AVSearchObj()
{
	this.Sequence = "";
	this.VendorID = "";
	this.NetworkID = "";
	this.Selected = false;	
}

function VendorSearch()
{
	this.VID = "";
	this.Selected = false;
}

function NetworkSearch()
{
	this.NetID = "";
	this.Selected = false;
}

var VendorSearchObj = new VendorSearch();
var NetObj = new NetworkSearch();
var AccVendObj = new AVSearchObj();
var lAlreadySaved = false;

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
<%
if lIsEdit then
%>
	document.all.TxtSequence.value = <%=oRS("SEQUENCE")%>
<%
	if not isNull(oRS.Fields("NETWORK_ID").Value) then
%>	
		document.all.Network_ID.value = <%=oRS("NETWORK_ID")%>
<%
	end if
	if not isNull(oRS.Fields("VENDOR_ID").Value) then
%>	
		document.all.Vendor_ID.value = <%=oRS("VENDOR_ID")%>
<%
	end if
%>			
	document.all.TxtContMeth.value = <%=oRS("CONTACT_METHOD_ID")%>
<%
	oRS.close
	set oRS = nothing
end if
%>	
End Sub

sub AttachVendor()

	VendorSearchObj.Selected = false
	strURL = "..\Vendors\VendorMaintenance.asp?SECURITYPRIV=FNSD_NETWORKS&CONTAINERTYPE=MODAL"
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	showModalDialog strURL, VendorSearchObj, "dialogWidth=600px; dialogHeight=550px; center=yes"
	If VendorSearchObj.Selected then
		document.all.Vendor_ID.value = VendorSearchObj.VID
	end if
end sub

sub AttachNet()
	NetObj.Selected = false
	strURL = "..\Vendors\NetworkMaintenance.asp?SECURITYPRIV=FNSD_NETWORKS&CONTAINERTYPE=MODAL"
	showModalDialog  strURL, NetObj, "dialogWidth=450px; dialogHeight=550px; center=yes"
	If NetObj.Selected then
		document.all.Network_ID.value = NetObj.NetID
	end if
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
	dim cErrMsg
	
	ValidateScreenData = true
	if document.all.TxtSequence.value = "" then
		cErrMsg = "Sequence is a required field." & chr(10) & chr(13)
	end if
	if cint(document.all.Network_ID.value) = 0 and cint(document.all.Vendor_ID.value) = 0 then
		cErrMsg = cErrMsg & "A Vendor ID or a Network ID must be supplied."
	end if
	if document.all.TxtContMeth.value = "" then
		cErrMsg = "Contact method is a required field." & chr(10) & chr(13)
	end if
	if len(cErrMsg) <> 0 then
		msgbox cErrMsg 
		ValidateScreenData = false
	End If
End Function

Function ExeSave
	if lAlreadySaved then
		exit function
	end if
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If
	
	document.all.VID.value = document.all.Vendor_ID.value
	document.all.NID.value = document.all.Network_ID.value
	document.all.SEQ.value = document.all.TxtSequence.value
	document.all.CM.value = document.all.TxtContMeth.value
	
	if ValidateScreenData = false then 
		ExeSave = false
		exit function
	end if
	lAlreadySaved = true
	document.body.setAttribute "ScreenDirty", "NO"
<%
if lIsEdit then
%>	
	document.all.ACTION.value = "EDIT"
	document.all.AVID.value = <%=Request.Querystring("AVID")%>
<%
end if
%>	
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
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>" IsThisRequired="<%=isRequired%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Add Vendor or Network to Account&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="AccVendorAddSave.asp" TARGET="hiddenPage">
<input type="hidden" NAME="VID">
<input type="hidden" NAME="NID">
<input type="hidden" NAME="SEQ">
<input type="hidden" NAME="LOB" VALUE='<%=Request.QueryString("LOB")%>'>
<input type="hidden" NAME="ST" VALUE='<%=Request.QueryString("ST")%>'>
<input type="hidden" NAME="AHSID" VALUE='<%=Request.QueryString("AHSID")%>'>
<input type="hidden" NAME="CM">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="ACTION">
<input type="hidden" NAME="AVID">
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

<table CELLPADDING=0 class="LABEL">
<tr>
<td><b>LOB:</b>&nbsp;<span id="spanAHSID"><%=cLOB%></span></td>
<td>&nbsp</td>
<td><b>Serv. Type:</b>&nbsp;<span id="spanAHSID"><%=cServiceType%></span></td>
</tr> 
<tr>
<td>&nbsp</td>
</tr>
<tr>
&nbsp
</tr>
</table>
<table>
<tr>
<TD CLASS=LABEL>Sequence:<BR><INPUT TYPE=TEXT SIZE=4 NAME=TxtSequence CLASS=LABEL></TD>
<td CLASS="LABEL">Vendor ID:<br>
<img SRC="../Images/Attach.gif" align=absmiddle ID="BtnATTACHVendor" STYLE="CURSOR:HAND" ALT="Attach Vendor" OnClick="AttachVendor()" WIDTH="16" HEIGHT="16">&nbsp
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="Vendor_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="0"></td>
<td CLASS="LABEL">&nbsp</td>
<td CLASS="LABEL">OR</td>
<td CLASS="LABEL">&nbsp</td>
<td CLASS="LABEL">Network ID:<br>
<img SRC="../Images/Attach.gif"  align=absmiddle ID="BtnATTACHNetwork" STYLE="CURSOR:HAND" ALT="Attach Network" OnClick="AttachNet()" WIDTH="16" HEIGHT="16">
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="Network_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="0"></td>
<td CLASS="LABEL">&nbsp</td>
<td CLASS="LABEL">Contact method:<br>
	<select ID="TxtContMeth" CLASS="LABEL" ScrnBtn="TRUE">
	<option VALUE>
	<%
	cSQL = "SELECT * FROM CONTACT_METHOD WHERE NAME IS NOT NULL"
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
	<option VALUE="<%=oRS("CONTACT_METHOD_ID")%>"><%= oRS("NAME") %>
	<%
	oRS.MoveNext
	Loop
	oRS.CLose
	oConn.close
	set oConn = nothing
	set oRS = nothing
	%>
</td>
</tr>
</table>
</form>
</body>
</html>


