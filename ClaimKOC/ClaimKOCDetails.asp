<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\validate.inc"-->

<%	Response.Expires = 0 
	Response.Buffer = true
		
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Claim Key Office Code Type Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language=javascript>

function CRoutingPlanSearchObj()
{
	this.RPID = "";
	this.RPDesc = "";
	this.Selected = false;
}

function CBranchSearchObj()
{
	this.BID = "";
	this.BIDOfficeName = "";
	this.BNUM = "";
	this.Selected = false;
}

var RoutingPlanSearchObj = new CRoutingPlanSearchObj();
var BranchSearchObj = new CBranchSearchObj();

var g_StatusInfoAvailable = false;
</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Function AttachBranch (ID, SPANID, strTITLE)
	BID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	BranchSearchObj.BID = BID
	BranchSearchObj.BIDOfficeName = ""
	BranchSearchObj.BNUM = ""
	BranchSearchObj.Selected = false

	If BID = "" Then BID = "NEW"
	
	If BID = "NEW" And MODE = "RO" Then
		MsgBox "No branch currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Branch\BranchMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CLAIM_ASSIGNMENT&SELECTONLY=TRUE&BranchTypeFilter=CLAIMHANDLING&BID=" & BID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  , BranchSearchObj ,"dialogWidth:550px;dialogHeight:450px;center"

	'if Selected=true update everything, otherwise if BID is the same, update text in case of save
	If BranchSearchObj.Selected = true Then
		If BranchSearchObj.BID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = BranchSearchObj.BID
		end if
		UpdateBranchText(SPANID)
	End If

End Function

Sub UpdateBranchText (SPANID)
	SPANID.innertext = BranchSearchObj.BNUM
	SPANID.title = BranchSearchObj.BNUM
End Sub

Function AttachRoutingPlan (ID, SPANID, strTITLE)
	RPID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	RoutingPlanSearchObj.RPID = RPID
	RoutingPlanSearchObj.RPDesc = ""
	RoutingPlanSearchObj.Selected = false

	If RPID = "" Then RPID = "NEW"
	
	If RPID = "NEW" And MODE = "RO" Then
		MsgBox "No Routing Plan currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\RoutingPlan\RoutingPlanMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CLAIM_ASSIGNMENT&SELECTONLY=TRUE"
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog strURL  , RoutingPlanSearchObj ,"center;dialogWidth:780px"

	'if Selected=true update everything, otherwise if BID is the same, update text in case of save
	If RoutingPlanSearchObj.Selected = true Then
		If RoutingPlanSearchObj.RPID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = RoutingPlanSearchObj.RPID
		end if
		UpdateRoutingPlanText(SPANID)
	End If

End Function

Sub UpdateRoutingPlanText (SPANID)
	SPANID.innertext = RoutingPlanSearchObj.RPDesc
	SPANID.title = RoutingPlanSearchObj.RPDesc
End Sub


Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub PostTo(strURL)
	FrmDetails.action = "ClaimKOCSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateKOCID(inKOCID)
	document.all.KOCID.value = inKOCID
	document.all.spanKOCID.innerText = inKOCID
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

Function GetKOCID
	if document.all.KOCID.value <> "NEW" then
		GetKOCID = document.all.KOCID.value
	else
		GetKOCID = ""
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
	'If  document.all.AHSID_ID.innerText = "" then
	'	MsgBox "A.H. Step ID is a required field.",0,"FNSNetDesigner"
	'	ValidateScreenData = false
	'	exit Function
	'end if
	
	
	
	if Len(document.all.TxtKOC.value) = 0 then
		MsgBox "'Claim KOC' is required!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	if not IsNumber(document.all.TxtSequence.value) then
		MsgBox "'Sequence' is required. Must be a number!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	if not IsNumber(document.all.TxtMinimum.value) then
		MsgBox "'Minimum' is required. Must be a number!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if

	if not IsNumber(document.all.TxtMaximum.value) then
		MsgBox "'Maximum' is required. Must be a number!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if

	if document.all.TxtMinimum.value > document.all.TxtMaximum.value then
		MsgBox "'Maximum' should be greater than 'Minimum'!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	if not IsNumber(document.all.TxtNext.value) then
		MsgBox "'Next' must be a number!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	if not IsNumber(document.all.TxtNotified.value) then
		MsgBox "'Notified' is required. Must be a number!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if

	if not IsNumber(document.all.TxtLength.value) then
		MsgBox "'Length' is required. Must be a number!",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if

	if Len(document.all.TxtWarningPercent.value) > 0 then
		if not IsNumber(document.all.TxtWarningPercent.value) then
			MsgBox "'Warning %' must be a number!",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
		if not CheckLenRange(document.all.TxtWarningPercent.value, 1, 2) then
			MsgBox "'Warning %' is invalid number!",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	end if
	
	ValidateScreenData = true
End Function

Function InEditMode
	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		InEditMode = false
	End If
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.KOCID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.KOCID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.KOCID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.KOCID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "CLAIM_KOC_ID" & Chr(129) & document.all.KOCID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BRANCH_NUMBER" & Chr(129) & document.all.BRANCH_NUMBER.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CLAIM_KOC" & Chr(129) & document.all.TxtKOC.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ROUTING_PLAN_ID" & Chr(129) & document.all.ROUTING_PLAN_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQ"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NEXT"& Chr(129) & document.all.TxtNext.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NOTIFIED_NUM" & Chr(129) & document.all.TxtNotified.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NOTIFY_EVERY" & Chr(129) & document.all.TxtNotifyEvery.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LENGTH"& Chr(129) & document.all.TxtLength.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "WARNINGPERCENT" & Chr(129) & document.all.TxtWarningPercent.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MINIMUM"& Chr(129) & document.all.TxtMinimum.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAXIMUM"& Chr(129) & document.all.TxtMaximum.value & Chr(129) & "1" & Chr(128)
		if document.all.ChkActiveFlg.checked = True then
			sResult = sResult & "ACTIVE_FLG"& Chr(129) & "Y"  & Chr(129) & "1" & Chr(128)
		else 
			sResult = sResult & "ACTIVE_FLG"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "ClaimKOCSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
			
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
<!--#include file="..\lib\Help.asp"-->
</script>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Claim Key Office Code Type Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="ClaimKOCSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchKOCID" value="<%=Request.QueryString("SearchKOCID")%>">
<input type="hidden" name="SearchBID" value="<%=Request.QueryString("SearchBID")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchBNUM" value="<%=Request.QueryString("SearchBNUM")%>">
<input type="hidden" name="SearchKOC" value="<%=Request.QueryString("SearchKOC")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="KOCID" value="<%=Request.QueryString("KOCID")%>">

<%	
Dim KOCID
KOCID	= CStr(Request.QueryString("KOCID"))
If KOCID <> "" Then
	If KOCID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT CLAIM_KOC_ASSIGNMENT.*, BRANCH.OFFICE_NAME, BRANCH.BRANCH_NUMBER, BRANCH.BRANCH_ID, ROUTING_PLAN.DESCRIPTION FROM " &_
				"CLAIM_KOC_ASSIGNMENT, BRANCH, ROUTING_PLAN WHERE " &_
				"CLAIM_KOC_ASSIGNMENT.BRANCH_NUMBER = BRANCH.BRANCH_NUMBER(+) AND " &_
				"CLAIM_KOC_ASSIGNMENT.ROUTING_PLAN_ID = ROUTING_PLAN.ROUTING_PLAN_ID(+) AND " &_
				"CLAIM_KOC_ID = " & KOCID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSBRANCH_ID = RS("BRANCH_ID")			
			RSBRANCH_OFFICE_NAME = RS("OFFICE_NAME")
			RSBRANCH_NUMBER = RS("BRANCH_NUMBER")
			RSKOC = RS("CLAIM_KOC")
			RSSEQUENCE = RS("SEQ")
			RSNEXT = RS("NEXT")
			RSLENGTH = RS("LENGTH")
			RSMINIMUM = RS("MINIMUM")
			RSMAXIMUM = RS("MAXIMUM")
			RSWARNINGPERCENT = RS("WARNINGPERCENT")
			RSACTIVE_FLG = RS("ACTIVE_FLG")
			RSNOTIFY_EVERY = RS("NOTIFY_EVERY")
			RSNOTIFIED_NUM = RS("NOTIFIED_NUM")
			RSRPID	= RS("ROUTING_PLAN_ID")
			RSRPDESC = RS("DESCRIPTION")
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	else
		RSNEXT  = "1"
		RSNOTIFIED_NUM = "0"
	End If

%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" >
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td colspan=2>Claim Key Office Code Type ID:&nbsp;<span id="spanKOCID"><%=Request.QueryString("KOCID")%></span></td></tr>
</table> 

<table LANGUAGE="JScript" class="Label" ONDRAGSTART="return false;">
<tr>
	<td>
	<IMG NAME=BtnAttachBranch STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Branch" ONCLICK="VBScript::AttachBranch BRANCH_ID, BRANCH_NUMBER,''">
	<IMG NAME=BtnDetachBranch STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Branch" OnClick="VBScript::Detach BRANCH_ID, BRANCH_NUMBER">
	</td>
	<td nowrap  WIDTH=150>Key Office Code:&nbsp<SPAN ID=BRANCH_NUMBER CLASS=LABEL TITLE="<%=RSBRANCH_NUMBER%>" ><%=RSBRANCH_NUMBER%></SPAN></td>
	<td>Branch ID:<span ID=BRANCH_ID><%=RSBRANCH_ID%></span></td>
</tr>


</table>


<table CLASS="LABEL" >
<tr>
	<td>Sub KOC<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtKOC" VALUE="<%=RSKOC%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text4"></input></td>
	<td>Sequence<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Length<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtLength" VALUE="<%=RSLENGTH%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
</tr>
<tr>
	<td>Minimum<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtMinimum" VALUE="<%=RSMINIMUM%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Maximum<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtMaximum" VALUE="<%=RSMAXIMUM%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
	<td>Warning %<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtWarningPercent" VALUE="<%=RSWARNINGPERCENT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></input></td>
</tr>
<tr>
	<td>Notify Every<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtNotifyEvery" VALUE="<%=RSNOTIFY_EVERY%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text3"></input></td>
	<td align=right valign=bottom>
		<IMG NAME=BtnAttachRoutingPlan STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach RoutingPlan" ONCLICK="VBScript::AttachRoutingPlan ROUTING_PLAN_ID, ROUTING_PLAN_DESC,''">
		<IMG NAME=BtnDetachRoutingPlan STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach RoutingPlan" OnClick="VBScript::Detach ROUTING_PLAN_ID, ROUTING_PLAN_DESC">
	</td>
	<td nowrap  WIDTH=150 align=left valign=bottom>Routing Plan ID:&nbsp<SPAN ID=ROUTING_PLAN_ID CLASS=LABEL TITLE="<%=RSRPID%>" ><%=RSRPID%></SPAN></td>
	<td valign=bottom>Routing Plan Description:&nbsp<span ID=ROUTING_PLAN_DESC><%=RSRPDESC%></span></td>
</tr>
<tr>
	<td>Next Counter<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtNext" VALUE="<%=RSNEXT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text2"></input></td>
	<td>Notified<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtNotified" VALUE="<%=RSNOTIFIED_NUM%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text1"></input></td>
	<td align=center valign=bottom>Active?<input ScrnBtn="TRUE" CLASS="LABEL" TYPE="CHECKBOX" NAME="ChkActiveFlg"  <% If CStr(RSACTIVE_FLG) = "Y" Then Response.Write("CHECKED")%> ONCLICK="VBScript::Control_OnChange"></input></td>
</tr>
</table> 

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Claim Key Office Code Type selected.
</div>


<% End If %>

</form>
</body>
</html>


