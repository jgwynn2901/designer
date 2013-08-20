<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%
	Response.Expires=0 
	
PROPID	= CStr(Request.QueryString("PROPID"))
RSPOLICY_ID = Request.QueryString("PID")
	
If PROPID <> "" Then
	If PROPID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM PROPERTY WHERE PROPERTY_ID = " & PROPID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSPROPERTY_ID = RS("PROPERTY_ID")
			RSPOLICY_ID = RS("POLICY_ID")
			RSPROPERTY_DESCRIPTION = ReplaceQuotesInText(RS("PROPERTY_DESCRIPTION"))
			RSPROPERTY_LOCATION_DESCRIPTION = ReplaceQuotesInText(RS("PROPERTY_LOCATION_DESCRIPTION"))
			RSADDRESS1 = ReplaceQuotesInText(RS("ADDRESS1"))
			RSADDRESS2 = ReplaceQuotesInText(RS("ADDRESS2"))
			RSCITY = RS("CITY")
			RSSTATE = RS("STATE")
			RSZIP = RS("ZIP")
			RSPHONE = RS("PHONE")
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
<title>Property Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JavaScript">
function CCoverageSearchObj()
{
	this.COVID = "";
	this.Selected = false;	
	this.Saved = false;	
}
function CPolicySearchObj()
{
	this.PID = "";
	this.Selected = "";
}
var PolicySearchObj = new CPolicySearchObj();
var CoverageSearchObj = new CCoverageSearchObj();

function EditCoverage()
{
	if (document.body.getAttribute("ScreenMode") == "RO") return;

	PROPID = document.all.PROPID.value;
	if ((PROPID == "") || (PROPID == "NEW")) return;


	CVID = document.frames("CoverageFrame").GetSelectedCVID()
	if (CVID != "")
	{
<%If HasModifyPrivilege("FNSD_COVERAGE","") <> True Then  %>		
		alert("You do not have the appropriate security privileges to edit coverage.");
		return;
<%End If %>		

		CoverageSearchObj.Saved = false;
		window.showModalDialog ("CoverageMaintenance.asp?SECURITYPRIV=FNSD_PROPERTY&CONTAINERTYPE=MODAL&DETAILONLY=TRUE&COVID=" + CVID, CoverageSearchObj, "center")
		if (CoverageSearchObj.Saved == true)
			RefreshCoverage();
	}
	else
		alert ("Please choose a coverage to edit");
	
}
function NewCoverage() 
{

	if (document.body.getAttribute("ScreenMode") == "RO") return;

	PROPID = document.all.PROPID.value;
	if ((PROPID == "") || (PROPID == "NEW")) return;

<%If HasAddPrivilege("FNSD_COVERAGE","FNSD_PROPERTY") <> True Then  %>		
		alert("You do not have the appropriate security privileges to add coverage.");
		return;
<%End If %>		

	CoverageSearchObj.Saved = false;
	window.showModalDialog ("CoverageMaintenance.asp?SECURITYPRIV=FNSD_PROPERTY&CONTAINERTYPE=MODAL&DETAILONLY=TRUE&COVID=NEW&PROPID=" + PROPID, CoverageSearchObj, "center")
	if (CoverageSearchObj.Saved == true)
		RefreshCoverage();
}
function RemoveCoverage() 
{
	if (document.body.getAttribute("ScreenMode") == "RO") return;

	PROPID = document.all.PROPID.value;
	if ((PROPID == "") || (PROPID == "NEW")) return;

	CVID = document.frames("CoverageFrame").GetSelectedCVID()
	if (CVID != "")
	{
<%If HasDeletePrivilege("FNSD_COVERAGE","FNSD_PROPERTY") <> True Then  %>		
		alert("You do not have the appropriate security privileges to delete coverage.");
		return;
<%End If %>		
		parent.frames("hiddenPage").location.href = "deletecoverage.asp?COVID=" + CVID
		RefreshCoverage();
	}
	else
		alert ("Please choose a coverage to delete");
}

function RefreshCoverage()
{
	PROPID = document.all.PROPID.value;
	document.frames("CoverageFrame").location.href = "PolicyDetailsCoverage.asp?PROPID=" + PROPID
}	

</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	else 
		if PROPID <> "" then %>
		document.all.TxtSTATE.Value = "<%= RSSTATE %>";
<%		end if	
	end if 
%>

if (document.all.CoverageFrame != null)
		document.all.CoverageFrame.style.height = .2 * document.body.clientHeight;
if (document.all.fldSet2 != null)
		document.all.fldSet2.style.height = .2 * document.body.clientHeight;

</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmDetails.action = "PropertySearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdatePROPID(inPROPID)
	document.all.PROPID.value = inPROPID
	document.all.spanPROPID.innerText = inPROPID
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

Function GetPROPID
	if document.all.PROPID.value <> "NEW" then
		GetPROPID = document.all.PROPID.value
	else
		GetPROPID = ""
	end if 
End Function

Function GetPROPIDName
	GetPROPIDName = document.all.TxtName.value
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
	If  document.all.TxtPOLICY_ID.value = "" then
		MsgBox "Policy ID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
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
	
	if document.all.PROPID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.PROPID.value = "NEW"
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
	
	if document.all.PROPID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.PROPID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "PROPERTY_ID"& Chr(129) & document.all.PROPID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "POLICY_ID"& Chr(129) & document.all.TxtPOLICY_ID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "PROPERTY_DESCRIPTION"& Chr(129) & document.all.TxtPROPERTY_DESCRIPTION.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PROPERTY_LOCATION_DESCRIPTION"& Chr(129) & document.all.TxtPROPERTY_LOCATION_DESCRIPTION.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS1"& Chr(129) & document.all.TxtADDRESS1.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS2"& Chr(129) & document.all.TxtADDRESS2.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY"& Chr(129) & document.all.TxtCITY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE"& Chr(129) & document.all.TxtSTATE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIP"& Chr(129) & document.all.TxtZIP.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPHONE.value & Chr(129) & "1" & Chr(128)
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

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Function AttachPolicy (ID)
	PID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	PolicySearchObj.PID = PID
	PolicySearchObj.Selected = false

	If PID = "" Then PID = "NEW"
	
	If PID = "NEW" And MODE = "RO" Then
		MsgBox "No policy currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Policy\PolicyMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_PROPERTY&SELECTONLY=TRUE&PID=" & PID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,PolicySearchObj ,"center"

	'if Selected=true update everything
	If PolicySearchObj.Selected = true Then
		If PolicySearchObj.PID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = PolicySearchObj.PID
		end if
	End If

End Function
<!--#include file="..\lib\Help.asp"-->
</script>
<script LANGUAGE="JavaScript" FOR="CoverageBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "EDITBUTTONCLICK":
			EditCoverage();
			break;

		case "NEWBUTTONCLICK":
			NewCoverage();
			break;

		case "REMOVEBUTTONCLICK":
			RemoveCoverage();
			break;
		default:
			break;
	}

</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Property Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="PropertySave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" name="SearchPROPID" value="<%=Request.QueryString("SearchPROPID")%>">
<input type="hidden" name="SearchPOLICY_ID" value="<%=Request.QueryString("SearchVEHICLE_ID")%>">
<input type="hidden" name="SearchPROPERTY_DESCRIPTION" value="<%=Request.QueryString("SearchNAME_FIRST")%>">
<input type="hidden" name="SearchADDRESS" value="<%=Request.QueryString("SearchADDRESS")%>">
<input type="hidden" name="SearchCITY" value="<%=Request.QueryString("SearchCITY")%>">
<input type="hidden" name="SearchSTATE" value="<%=Request.QueryString("SearchSTATE")%>">
<input type="hidden" name="SearchZIP" value="<%=Request.QueryString("SearchZIP")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PROPID" value="<%=Request.QueryString("PROPID")%>">

<%	
If PROPID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<!--<td WIDTH="14"><img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count"></td><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">:<span id="SpanSharedCount"><%=SharedCount%></span></td>-->
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL">
	<tr>
	<td COLSPAN="5" CLASS="LABEL">Property ID:&nbsp;<span id="spanPROPID"><%=Request.QueryString("PROPID")%></span></td>
	</tr>
	<tr>
	<td VALIGN="BOTTOM"><img NAME="BtnAttachPolicy" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Policy" ONCLICK="VBScript::AttachPolicy TxtPOLICY_ID">
	<td CLASS="LABEL" COLSPAN="2">Policy ID:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" READONLY STYLE="BACKGROUND-COLOR:SILVER" MAXLENGTH="10" TYPE="TEXT" NAME="TxtPOLICY_ID" VALUE="<%=RSPOLICY_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
	<table>
	<tr>	
	<td CLASS="LABEL">Property Description:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="255" size="90" TYPE="TEXT" NAME="TxtPROPERTY_DESCRIPTION" VALUE="<%= RSPROPERTY_DESCRIPTION %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Property Location Description:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="255" size="90" TYPE="TEXT" NAME="TxtPROPERTY_LOCATION_DESCRIPTION" VALUE="<%=RSPROPERTY_LOCATION_DESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
</table>
<table>
<tr>
	<td CLASS="LABEL">Address 1:<br><input ScrnInput="TRUE" size="98" TYPE="TEXT" MAXLENGTH="80" NAME="TxtADDRESS1" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSADDRESS1%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Address 2:<br><input ScrnInput="TRUE" size="98" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" NAME="TxtADDRESS2" VALUE="<%=RSADDRESS2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
</table>
<table>
<tr>
	<td CLASS="LABEL">City:<br><input ScrnInput="TRUE" size="40" TYPE="TEXT" MAXLENGTH="40" NAME="TxtCITY" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSCITY%>"></td>
	<td CLASS="LABEL">State:<br>
	<select ScrnBtn="TRUE" CLASS="LABEL" NAME="TxtSTATE" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	<!--#include file="..\lib\States.asp"-->
	</select>
	<td CLASS="LABEL">Zip:<br><input ScrnInput="TRUE" size="9" CLASS="LABEL" MAXLENGTH="9" TYPE="TEXT" NAME="TxtZip" VALUE="<%=RSZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS="LABEL">Phone Number:<br><input ScrnInput="TRUE" size="15" TYPE="TEXT" MAXLENGTH="14" NAME="TxtPHONE" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSPHONE%>"></td>
	</tr>
	</table>

<table CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Coverage</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="175" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset ID="fldSet2" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100';width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&amp;HIDEATTACH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="CoverageBtnControl" type="text/x-scriptlet"></object>
<iframe FRAMEBORDER="0" name="CoverageFrame" width="100%" height="100" src="PolicyDetailsCoverage.asp?<%=Request.QueryString%>"></iframe>
</fieldset>
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No property selected.
</div>
<% End If %>

</form>
</body>
</html>


