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
<title>User's Groups</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function type_groupsSearchObj()
{
	this.Selected = false;
}
var groupsSearchObj = new type_groupsSearchObj();
var g_StatusInfoAvailable = false;

</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 100;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 100;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmGroups.action = strURL
	FrmGroups.method = "GET"
	FrmGroups.target = "_parent"	
	FrmGroups.submit
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

Function GetSelectedGID
	GetSelectedGID = document.frames("DataFrame").GetSelectedGID
End Function

Sub Refresh
	UID = document.all.UID.value
	document.all.tags("IFRAME").item("DataFrame").src = "UserGroupsData.asp?UID=" & UID
End Sub

Function InEditMode
	InEditMode = true
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "This screen is read only.",0,"FNSNetDesigner"
		InEditMode = false
	End If

End Function

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
	End If		
End Sub

Sub ExeButtonsAttach
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.UID.value = "" Or document.all.UID.value = "NEW" Then
		Exit Sub
	End If

	groupsSearchObj.Selected = false

	strURL = "../Groups/GroupMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_USERS"
	showModalDialog  strURL,groupsSearchObj ,"center"

	If groupsSearchObj.Selected = true Then	
		GID = groupsSearchObj.GID
		If GID <> "" Then
			sResult = "(" & document.all.UID.value
			sResult = sResult & " ," & GID & ")"
			document.all.TxtSaveData.Value = sResult
			document.all.TxtAction.Value = "INSERT"
			FrmGroups.action = "UserGroupSave.asp"
			FrmGroups.method = "POST"
			FrmGroups.target = "hiddenPage"	
			FrmGroups.submit
			Refresh
		End If
	End If
End Sub

Sub ExeButtonsRemove
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.UID.value = "" Or document.all.UID.value = "NEW" Then
		Exit Sub
	End If

	dim GID, sResult
	GID = GetSelectedGID

	If GID <> "" Then
		sResult = "USER_ID = " &  document.all.UID.value
		sResult = sResult & " AND GROUP_ID = " & GID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmGroups.action = "UserGroupSave.asp"
		FrmGroups.method = "POST"
		FrmGroups.target = "hiddenPage"	
		FrmGroups.submit
		Refresh
	Else
		MsgBox "Please select a group to detach the current user from.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub
</script>
<!--#include file="..\lib\UserBtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Groups</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmGroups" ID="Form1">
<input TYPE="HIDDEN" NAME="TxtSaveData" ID="Hidden1">
<input TYPE="HIDDEN" NAME="TxtAction" ID="Hidden2">
<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchUID" value="<%=Request.QueryString("SearchUID")%>" ID="Hidden3">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>" ID="Hidden4">
<input type="hidden" name="SearchSite" value="<%=Request.QueryString("SearchSite")%>" ID="Hidden5">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>" ID="Hidden6">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" ID="Hidden7">
<input type="hidden" NAME="UID" value="<%=Request.QueryString("UID")%>" ID="Hidden8">

<%	
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
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" ID="Table3">
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" ID="Table4">
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
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="UserGroupsData.asp?<%=Request.QueryString%>">
</fieldset>
</form>
</body>
</html>
