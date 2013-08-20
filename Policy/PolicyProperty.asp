<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->

<%	Response.Expires=0 
	Response.Buffer = true

	Dim PID
	PID	= CStr(Request.QueryString("PID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Property</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CPropertySearchObj()
{
	this.Selected = false;
	this.Saved = false;

}

var PropertySearchObj = new CPropertySearchObj();
var g_StatusInfoAvailable = false;

</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	'If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	//SetScreenFieldsReadOnly(true,"DISABLED");
<%	'End If %>
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 100;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 100;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmProperty.action = strURL
	FrmProperty.method = "GET"
	FrmProperty.target = "_parent"	
	FrmProperty.submit
End Sub

Sub UpdatePID(inPID)
	document.all.PID.value = inPID
	document.all.spanPID.innerText = inPID
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

Function GetPID
	GetPID = document.all.PID.value
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

Function GetSelectedPROPID
	GetSelectedPROPID = document.frames("DataFrame").GetSelectedPROPID
End Function


Sub Refresh
	PID = document.all.PID.value
	document.all.tags("IFRAME").item("DataFrame").src = "PolicyPropertyData.asp?PID=" & PID
End Sub

Sub ExeButtonsNew
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.PID.value = "" Or document.all.PID.value = "NEW" Then
		Exit Sub
	End If

<%If HasAddPrivilege("FNSD_PROPERTY","FNSD_POLICY") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to add properties.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

	PropertySearchObj.Saved = false

	dim MODE, PID
	MODE = document.body.getAttribute("ScreenMode")
	PID = document.all.PID.value
	strURL = "PropertyMaintenance.asp?SECURITYPRIV=FNSD_POLICY&PID=" & PID & "&PROPID=NEW&DETAILONLY=TRUE"
	showModalDialog  strURL,PropertySearchObj ,"center"

	If PropertySearchObj.Saved = true Then Refresh
End Sub

Sub ExeButtonsEdit
	If Not InEditMode Then
		Exit Sub
	End If

	PID = document.all.PID.value

	If PID = "" Or PID = "NEW" Then Exit Sub

	dim PROPID
	PROPID = GetSelectedPROPID
	
	If PROPID <> "" Then
	
<%If HasModifyPrivilege("FNSD_PROPERTY","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to edit properties.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		
	
		PropertySearchObj.Saved = false
		dim MODE
		MODE = document.body.getAttribute("ScreenMode")
		strURL = "PropertyMaintenance.asp?SECURITYPRIV=FNSD_POLICY&PID=" & PID & "&PROPID=" & PROPID & "&DETAILONLY=TRUE"
		showModalDialog  strURL,PropertySearchObj ,"center"
		If PropertySearchObj.Saved = true Then Refresh
	Else
		MsgBox "Please select a property to edit.", 0, "FNSNet Designer"		
	End If
End Sub

Sub ExeButtonsRemove
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.PID.value = "" Or document.all.PID.value = "NEW" Then
		Exit Sub
	End If

	dim PROPID, sResult
	PROPID = GetSelectedPROPID
	
	If PROPID <> "" Then
<%If HasDeletePrivilege("FNSD_PROPERTY","FNSD_POLICY") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to delete properties.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		
		sResult = sResult &  PROPID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmProperty.action = "PropertySave.asp"
		FrmProperty.method = "POST"
		FrmProperty.target = "hiddenPage"	
		FrmProperty.submit
		Refresh
	Else
		MsgBox "Please select a property to remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub

Function InEditMode
	InEditMode = true
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "This screen is read only.",0,"FNSNetDesigner"
		InEditMode = false
	End If

End Function

Function ExeCopy
	If Not InEditMode Then
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.PID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
		
	FrmProperty.action = "PolicyCopy.asp"
	FrmProperty.method = "GET"
	FrmProperty.target = "hiddenPage"	
	FrmProperty.submit
'	Refresh is done inside UserCopy.asp due to timing
	ExeCopy = true
End Function

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
	End If		
End Sub


</script>
<SCRIPT LANGUAGE="JavaScript" FOR="PropertyBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "REMOVEBUTTONCLICK":
				ExeButtonsRemove();
			break;
		case "EDITBUTTONCLICK":
				ExeButtonsEdit();
			break;
		case "NEWBUTTONCLICK":
				ExeButtonsNew();
			break;
		default:
			break;
	}
</SCRIPT>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Policy Property</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmProperty">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchPID" value="<%=Request.QueryString("SearchPID")%>">
<input type="hidden" name="SearchNumber" value="<%=Request.QueryString("SearchNumber")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchCarrier" value="<%=Request.QueryString("SearchCarrier")%>">
<input type="hidden" name="SearchAgent" value="<%=Request.QueryString("SearchAgent")%>">
<input type="hidden" name="SearchLOBCD" value="<%=Request.QueryString("SearchLOBCD")%>">
<input type="hidden" name="SearchSelfInsuredFlg" value="<%=Request.QueryString("SearchSelfInsuredFlg")%>">
<input type="hidden" name="SearchEffective" value="<%=Request.QueryString("SearchEffective")%>">
<input type="hidden" name="SearchOriginalEffective" value="<%=Request.QueryString("SearchOriginalEffective")%>">
<input type="hidden" name="SearchExpiration" value="<%=Request.QueryString("SearchExpiration")%>">
<input type="hidden" name="SearchCancellation" value="<%=Request.QueryString("SearchCancellation")%>">
<input type="hidden" name="SearchChange" value="<%=Request.QueryString("SearchChange")%>">
<input type="hidden" name="SearchLoad" value="<%=Request.QueryString("SearchLoad")%>">
<input type="hidden" name="SearchCompanyCode" value="<%=Request.QueryString("SearchCompanyCode")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PID" value="<%=Request.QueryString("PID")%>">

<%	

If PID <> "" Then
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
<tr>
<tr>
<tr>
<tr>
<tr>
<td>Policy ID:&nbsp;<span id="spanPID"><%=Request.QueryString("PID")%></span></td>
</tr>
<tr>
</table>

<span class="Label" ID=SPANDATA>Retrieving...</span>
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?NEWCAPTION=Add&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDEATTACH=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=PropertyBtnControl type=text/x-scriptlet></OBJECT>
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="PolicyPropertyData.asp?<%=Request.QueryString%>">
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No policy selected.
</div>


<% End If %>

</form>
</body>
</html>


