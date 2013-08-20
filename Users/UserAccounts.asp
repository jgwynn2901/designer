<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->

<%	Response.Expires=0 
	Response.Buffer = true

	Dim UID
	UID	= CStr(Request.QueryString("UID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Accounts</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();
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
	FrmAccounts.action = strURL
	FrmAccounts.method = "GET"
	FrmAccounts.target = "_parent"	
	FrmAccounts.submit
End Sub

Sub UpdateUID(inUID)
	document.all.UID.value = inUID
	document.all.spanUID.innerText = inUID
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

Function GetUID
	GetUID = document.all.UID.value
End Function

Function GetUIDName
	GetUIDName = document.all.spanName.innerText
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

Function GetSelectedAHSID
	GetSelectedAHSID = document.frames("DataFrame").GetSelectedAHSID
End Function


Sub ExeButtonsAttach
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.UID.value = "" Or document.all.UID.value = "NEW" Then
		Exit Sub
	End If

	AHSSearchObj.Selected = false

	dim MODE, UID
	MODE = document.body.getAttribute("ScreenMode")
	UID = document.all.UID.value
	strURL = "../AH/AHSMaintenance.asp?SEARCHONLY=TRUE&SELECTONLY=TRUE&SECURITY_PRIV=FNSD_USERS&OriginSource=USERS"
	showModalDialog  strURL,AHSSearchObj ,"center"

	If AHSSearchObj.Selected = true Then	
		If AHSSearchObj.AHSID <> "" Then
			sResult = "{call Designer.AddAccountUser('" & AHSSearchObj.AHSID & "', '" &_
					document.all.UID.value & "', {resultset  2, StatusMsg,StatusNum})}"
			document.all.TxtSaveData.Value = sResult
			document.all.TxtAction.Value = "INSERT"
			FrmAccounts.action = "UserAccountsSave.asp"
			FrmAccounts.method = "POST"
			FrmAccounts.target = "hiddenPage"	
			FrmAccounts.submit
			Refresh
		End If
	End If
End Sub


Sub Refresh
	UID = document.all.UID.value
	document.all.tags("IFRAME").item("DataFrame").src = "UserAccountsData.asp?UID=" & UID
End Sub

Sub ExeButtonsRemove
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.UID.value = "" Or document.all.UID.value = "NEW" Then
		Exit Sub
	End If

	dim AHSID, sResult
	AHSID = GetSelectedAHSID
	
	If AHSID <> "" Then
		sResult = "USER_ID = " &  document.all.UID.value
		sResult = sResult & " AND ACCNT_HRCY_STEP_ID = " & AHSID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmAccounts.action = "UserAccountsSave.asp"
		FrmAccounts.method = "POST"
		FrmAccounts.target = "hiddenPage"	
		FrmAccounts.submit
		Refresh
	Else
		MsgBox "Please select an account to detach the current user from.", 0, "FNSNet Designer"
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
	
	If document.all.UID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
		
	FrmAccounts.action = "UserCopy.asp"
	FrmAccounts.method = "GET"
	FrmAccounts.target = "hiddenPage"	
	FrmAccounts.submit
	'Refresh is done inside UserCopy.asp due to timing
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
<!--#include file="..\lib\UserBtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Accounts</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmAccounts">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchUID" value="<%=Request.QueryString("SearchUID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchSite" value="<%=Request.QueryString("SearchSite")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="UID" value="<%=Request.QueryString("UID")%>">

<%	

If UID <> "" Then
	If UID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT NAME FROM USERS_SITE_VIEW WHERE USER_ID = " & UID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSNAME = RS("NAME")
		End If

		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing

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
<tr>
<tr>
<tr>
<tr>
<tr>
<td>User ID:&nbsp;<span id="spanUID"><%=Request.QueryString("UID")%></span></td>
<td>&nbsp;&nbsp;</td>
<td>Name:&nbsp;<span id="spanName"><%=RSNAME%></span></td>
</tr>
<tr>
</table>

<span class="Label" ID=SPANDATA>Retrieving...</span>
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&REMOVECAPTION=Detach&HIDENEW=TRUE&HIDEEDIT=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=UserBtnControl type=text/x-scriptlet></OBJECT>
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="UserAccountsData.asp?<%=Request.QueryString%>">
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No user selected.
</div>


<% End If %>

</form>
</body>
</html>


