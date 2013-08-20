<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\security.inc"-->

<%
Response.Expires=0 
AccountTextLen = 30	

Dim cAccVendorID
dim cSQL, oRS, oConn, cAHSID
dim nST
	
cAccVendorID = trim(Request.QueryString("AVID"))
cAHSID = trim(Request.QueryString("AHSID"))
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
RSACCNT_HRCY_STEP_ID= cAHSID	
cSQL = "SELECT NAME FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID = " & cAHSID
Set oRS = oConn.Execute(cSQL)
If Not oRS.EOF then
	RSACCOUNT_NAME= ReplaceQuotesInText(oRS("NAME"))
end if
oRS.Close
If len(cAccVendorID) <> 0 Then
	If cAccVendorID <> "NEW" then
		cSQL = "SELECT * FROM ACCOUNT_VENDOR WHERE ACCOUNT_VENDOR_ID = " & cAccVendorID
		Set oRS = oConn.Execute(cSQL)
		If Not oRS.EOF then
			RS_LOB = oRS("LOB")
			RS_ST = oRS("SERVICE_TYPE_ID")
		end if
		oRS.Close
		Set oRS = Nothing
	end if	
End If
%>
<html>
<head>
<title>Account Vendors Detail</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language=jscript>
function AVSearchObj()
{
	this.Sequence = "";
	this.VendorID = "";
	this.NetworkID = "";
	this.Selected = false;	
}

var AccVendObj = new AVSearchObj();
var lVendorAdded = false;
var nSeq;
var nVendorID;
var nNetworkID;
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
dim cLOB

if document.all.DataFrame <> null then
	document.all.DataFrame.style.width = document.body.clientWidth - 175
end if

<%	
if Request.QueryString("MODE") = "RO" then 
%>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	
elseif len(cAccVendorID) <> 0 and cAccVendorID<>"NEW" then 
%>
	document.all.TxtLOB_CD.VAlue = "<%=RS_LOB%>"
	document.all.TxtServType.value = <%=RS_ST%>
<%
end if	
%>
End Sub

Sub UpdateAVID(inAVID)
	document.all.AVID.value = inAVID
	document.all.spanAVID.innerText = inAVID
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

Function GetAVID
	if document.all.AVID.value <> "NEW" then
		GetAVID = document.all.AVID.value
	else
		GetAVID = ""
	end if 
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
dim cErrmsg

If document.all.AHSID_ID.innerText = "" then
	cErrmsg = "Accnt Hrcy Step ID is a required field." & VbCrlf
end if
If document.all.TxtLOB_CD.value = "" then
	cErrmsg = cErrmsg & "Line of Business is a required field." & VbCrlf
end if

If document.all.TxtServType.value = "" then
	cErrmsg = cErrmsg & "Service Type is a required field." & VbCrlf
end if
If len(cErrmsg) = 0 Then
	ValidateScreenData = true
Else
	msgbox cErrmsg, 0, "FNSDesigner"
End If
End Function

Function ExeSave
	sResult = ""
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	if document.all.AVID.value = "" then
		ExeSave = false
		exit function
	end if
	
	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.AVID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "ACCOUNT_VENDOR_ID"& Chr(129) & document.all.AVID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "0" & Chr(128)
		sResult = sResult & "LOB"& Chr(129) & document.all.TxtLOB_CD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SERVICE_TYPE_ID"& Chr(129) & document.all.TxtServType.value & Chr(129) & "0" & Chr(128)
		If document.all.AVID.value = "NEW" then
			sResult = sResult & "CONTACT_METHOD_ID"& Chr(129) & "-1" & Chr(129) & "0" & Chr(128)
		end if
		document.all.TxtSaveData.Value = sResult
		document.all.LOB.Value = document.all.TxtLOB_CD.value
		document.all.ST.Value = document.all.TxtServType.value
		document.all.FrmDetails.Submit()
		bRet = true
	Else
		SpanStatus.innerHTML = "Nothing to Save"
	End If
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

Sub Refresh
	document.all.tags("IFRAME").item("DataFrame").src = "AccVendorDetailsData.asp?LOB=" & document.all.TxtLOB_CD.Value & "&ST=" & document.all.TxtServType.value & "&AHSID=<%=cAHSID%>&AVID=" & document.all.AVID.value 
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Sub AddAccVendor
<%If HasAddPrivilege("FNSD_ACC_VENDOR","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to add branch assignment rules.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

	AVID = document.all.AVID.value
	if AVID = "NEW" then
		msgbox "You have to save first, before adding Vendors or Networks."
		exit sub
	end if
	MODE = document.body.getAttribute("ScreenMode")
	
	AccVendObj.Selected = false
	strURL = "..\AccountVendorsAdd\AccVendorAddModal.asp?LOB=" & document.all.TxtLOB_CD.Value & "&ST=" & document.all.TxtServType.Value & "&<%=Request.querystring%>"

	showModalDialog strURL, AccVendObj, "dialogWidth=500px; dialogHeight=280px; center=yes"
	If AccVendObj.Selected Then	
		lVendorAdded = true
		nSeq = AccVendObj.Sequence
		nVendorID = AccVendObj.VendorID
		nNetworkID = AccVendObj.NetworkID
		Refresh
	end if
End Sub

Function GetSelectedAVID
	GetSelectedAVID = document.frames("DataFrame").GetSelectedAVID
End Function

Sub EditAccVendor
dim nAVID, sResult

	If document.all.AVID.value = "" Or document.all.AVID.value = "NEW" Then
		Exit Sub
	End If

<%If HasDeletePrivilege("FNSD_ACC_VENDOR","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to delete branch assignment rules.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

nAVID = GetSelectedAVID
If nAVID <> "" Then
	strURL = "..\AccountVendorsAdd\AccVendorAddModal.asp?EDIT=True&AVID=" & nAVID & "&LOB=" & document.all.TxtLOB_CD.Value & "&ST=" & document.all.TxtServType.Value
	AccVendObj.Selected = false
	showModalDialog strURL, AccVendObj, "dialogWidth=500px; dialogHeight=280px; center=yes"
	If AccVendObj.Selected Then	
		Refresh
	end if
Else
	MsgBox "Please select a Vendor or Network to edit.", 0, "FNSDesigner"		
End If
end sub


Sub DelAccVendor
dim nAVID, sResult

	If document.all.AVID.value = "" Or document.all.AVID.value = "NEW" Then
		Exit Sub
	End If

<%If HasDeletePrivilege("FNSD_ACC_VENDOR","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to delete branch assignment rules.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

nAVID = GetSelectedAVID
If nAVID <> "" Then
	sResult = nAVID
	document.all.TxtSaveData.Value = sResult
	document.all.TxtAction.Value = "DELETE"
	FrmDetails.action = "AccVendorSave.asp"
	FrmDetails.method = "POST"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit
	Refresh
Else
	MsgBox "Please select a Vendor or Network to Remove.", 0, "FNSDesigner"		
End If
end sub

<!--#include file="..\lib\Help.asp"-->
</script>
<!--#include file="..\lib\AVBtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Account Vendors Detail&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<form Name="FrmDetails" METHOD="POST" ACTION="AccVendorSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="AVID" value="<%=Request.QueryString("AVID")%>">
<input TYPE="HIDDEN" NAME="LOB">
<input TYPE="HIDDEN" NAME="ST">
<input TYPE="HIDDEN" NAME="AHS_ID" value="<%=Request.QueryString("AHSID")%>">
<%	

If cAccVendorID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<!--<td WIDTH="14"><img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count"></td><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">:<span id="SpanSharedCount"><%=SharedCount%></span></td>-->
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
<td>
<input ScrnBtn="TRUE" STYLE="DISPLAY:NONE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL">
<tr>
	<td width="305" nowrap>Account:&nbsp;<span ID="AHSID_TEXT" CLASS="LABEL" TITLE="<%=ReplaceQuotesInText(RSACCOUNT_NAME)%>"><%=TruncateText(RSACCOUNT_NAME,AccountTextLen)%></span></td>
	<td>A.H.S. ID:&nbsp;<span ID="AHSID_ID" CLASS="LABEL"><%=RSACCNT_HRCY_STEP_ID%></span></td>
	</tr>
</table>

<table class="LABEL">
	<tr>
	<td COLSPAN="5" CLASS="LABEL">Account Vendor ID:&nbsp;<span id="spanAVID"><%=Request.QueryString("AVID")%></span></td>
	</tr>
	</table>
	
	<table>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">LOB:<br>
	<!--<select NAME="TxtLOB_CD" CLASS="LABEL" ScrnBtn="TRUE" onchange="chkINF(this.options[this.selectedIndex].value)">-->
	<select NAME="TxtLOB_CD" CLASS="LABEL" ScrnBtn="TRUE" ONCHANGE="VBScript::Control_OnChange">
	<option VALUE>
	<%
	SQLST = "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
	Set RS = oConn.Execute(SQLST)
	Do While Not RS.EOF
	%>
	<option VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_CD") %>
	<%
	RS.MoveNext
	Loop
	RS.CLose
	%>
	</select></td>
	<td CLASS="LABEL" COLSPAN="4">Service Type:<br>
	<select ID="TxtServType" CLASS="LABEL" ScrnBtn="TRUE" ONCHANGE="VBScript::Control_OnChange">
	<option VALUE>
	<%
	SQLST = "SELECT * FROM SERVICE_TYPE WHERE TYPE IS NOT NULL"
	Set RS = oConn.Execute(SQLST)
	Do While Not RS.EOF
	%>
	<option VALUE="<%=RS("SERVICE_TYPE_ID")%>"><%= RS("TYPE") %>
	<%
	RS.MoveNext
	Loop
	RS.CLose
	oConn.close
	set RS = nothing
	set oConn = nothing
	%>
	</select></td>
	</tr>
	</table>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vendors</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset id="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?NEWCAPTION=Add&HIDEREFRESH=TRUE&amp;HIDEATTACH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE&amp;" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="AVBtnControl" type="text/x-scriptlet" VIEWASTEXT></object>
<iframe width="198%" height="185%" name="DataFrame" src="AccVendorDetailsData.asp?LOB=<%=RS_LOB%>&ST=<%=RS_ST%>&<%=Request.QueryString%>">
</fieldset>
	

<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Vendor selected.
</div>
<% End If %>
</form>
</body>
</html>


