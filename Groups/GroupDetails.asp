<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	Response.Expires=0 

	Dim GID
	GID	= CStr(Request.QueryString("GID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Group Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CUserSearchObj()
{
	this.UID = "";
	this.UIDName = "";
	this.Selected = false;
}
var UserSearchObj = new CUserSearchObj();
var g_StatusInfoAvailable = false;
</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 150;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 150;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub PostTo(strURL)
	FrmDetails.action = strURL
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateGID(inGID)
	document.all.GID.value = inGID
	document.all.spanGID.innerText = inGID
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

Function GetGID
	if document.all.GID.value <> "NEW" then
		GetGID = document.all.GID.value
	else
		GetGID = ""
	end if 
End Function

Function GetGIDName
	GetGIDName = document.all.TxtName.value
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
	If  document.all.TxtName.value = "" Then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		Exit Function
	End If

	ValidateScreenData = true
End Function

Function GetSelectedUID
	GetSelectedUID = document.frames("DataFrame").GetSelectedUID
End Function

Sub ExeButtonsAttach
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.GID.value = "" Or document.all.GID.value = "NEW" Then
		Exit Sub
	End If

	UserSearchObj.Selected = false

	strURL = "../Users/UserMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_GROUPS"
	showModalDialog  strURL,UserSearchObj ,"center"

	If UserSearchObj.Selected = true Then	
		UID = UserSearchObj.UID
		If UID <> "" Then
			sResult = "(" & UID
			sResult = sResult & " ," & document.all.GID.value & ")"
			document.all.TxtSaveData.Value = sResult
			document.all.TxtAction.Value = "INSERT"
			FrmDetails.action = "../Users/UserGroupSave.asp"
			FrmDetails.method = "POST"
			FrmDetails.target = "hiddenPage"	
			FrmDetails.submit
			Refresh
		End If
	End If
End Sub

Sub Refresh
	GID = document.all.GID.value
	document.all.tags("IFRAME").item("DataFrame").src = "GroupDetailsData.asp?GID=" & GID
End Sub

Sub ExeButtonsRemove
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.GID.value = "" Or document.all.GID.value = "NEW" Then
		Exit Sub
	End If

	dim UID, sResult
	UID = GetSelectedUID

	If UID <> "" Then
		sResult = "GROUP_ID = " &  document.all.GID.value
		sResult = sResult & " AND USER_ID = " & UID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmDetails.action = "../Users/UserGroupSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a user to detach from the current group.", 0, "FNSNet Designer"		
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
	
	If document.all.GID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
		
	FrmDetails.action = "GroupCopy.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit
'	Refresh is done inside GroupCopy.asp due to timing
	ExeCopy = true
End Function

Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.GID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.GID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		sResult = sResult & "GROUP_ID"& Chr(129) & document.all.GID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "GROUP_NM"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "GroupSave.asp"
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


</script>
<!--#include file="..\lib\GroupBtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Group Details</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="GroupSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchGID" value="<%=Request.QueryString("SearchGID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="GID" value="<%=Request.QueryString("GID")%>">

<%	

If GID <> "" Then
	If GID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM GROUPS WHERE GROUP_ID = " & GID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSNAME = ReplaceQuotesInText(RS("GROUP_NM"))
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing

	End If
%>
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

<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td>Group ID:&nbsp;<span id="spanGID"><%=Request.QueryString("GID")%></span></td></tr>
<tr>
<td>Name:<br><input ScrnInput="TRUE" MAXLENGTH="80" CLASS="LABEL" size="30" TYPE="TEXT" NAME="TxtName" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</table>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Users</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<span class="Label" ID=SPANDATA>Retrieving...</span>
<fieldset id="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&REMOVECAPTION=Detach&HIDEEDIT=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDENEW=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="GroupBtnControl" type="text/x-scriptlet"></object>
<iframe FRAMEBORDER="0" width="100%" height=0 name="DataFrame" src="GroupDetailsData.asp?<%=Request.QueryString%>">
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No group selected.
</div>


<% End If %>

</form>
</body>
</html>


