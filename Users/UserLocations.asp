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
<title>User Locations</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CAccessPermissionsSearchObj()
{
	this.Selected = false;
}
var AccessPermissionsSearchObj = new CAccessPermissionsSearchObj();
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
	FrmPermissions.action = strURL
	FrmPermissions.method = "GET"
	FrmPermissions.target = "_parent"	
	FrmPermissions.submit
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

Function GetSelectedACCID
	GetSelectedACCID = document.frames("DataFrame").GetSelectedACCID
End Function

Sub ExeButtonsNew()
	If Not InEditMode Then
		Exit Sub
	End If
     If document.all.UID.value = "" Or document.all.UID.value = "NEW" Then
		 Exit Sub
	 End If
'' If ACCID <> "" Then
	 AccessPermissionsSearchObj.Selected = false
	  dim MODE, UID, TITLE
	  MODE = document.body.getAttribute("ScreenMode")
	  UID = document.all.UID.value
	  TITLE = document.all.spanName.innerText & " USER" 
	  strURL = "AccessLocationsModal.asp?MODE=" & MODE & "&UID=" & UID & "&ACCID=NEW" 
	  showModalDialog  strURL,AccessPermissionsSearchObj ,"center"

	If AccessPermissionsSearchObj.Selected = true Then	Refresh
	
	''Else
	''	MsgBox "User location ID exist.", 0, "FNSNet Designer"		
	'''End If
	
End Sub



Sub ExeButtonsEdit()
     dim MODE, UID

	If Not InEditMode Then
		Exit Sub
	End If
  UID = document.all.UID.value
  ACCID = GetSelectedACCID
   If UID = "" Or UID = "NEW" Then Exit Sub
    If ACCID <> "" Then
          AccessPermissionsSearchObj.Selected = false
		  MODE = document.body.getAttribute("ScreenMode")
		  
		strURL = "AccessLocationsModal.asp?UID=" & UID & "&ACCID=" & ACCID & "&DETAILONLY=TRUE"
		showModalDialog strURL, AccessPermissionsSearchObj, "dialogWidth=500px; dialogHeight=610px; center=yes"
		If AccessPermissionsSearchObj.Selected Then 
			Refresh
		end if
	Else
		MsgBox "Please select a User location ID to edit.", 0, "FNSNet Designer"		
	End If

End Sub


Sub Refresh
	UID = document.all.UID.value
	document.all.tags("IFRAME").item("DataFrame").src = "UserLocationsData.asp?UID=" & UID
End Sub

Sub ExeButtonsRemove()
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.UID.value = "" Or document.all.UID.value = "NEW" Then
		Exit Sub
	End If

	dim ACCID, sResult
	ACCID = GetSelectedACCID
	
	If ACCID <> "" Then
		If Not document.frames("DataFrame").IsSelectedUserLevel() Then
			MsgBox "You may only remove user level access permissions.", 0, "FNSNet Designer"		
			Exit Sub
		End If

		sResult = sResult &  ACCID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmPermissions.action = "AccessLocationsSave.asp"
		FrmPermissions.method = "POST"
		FrmPermissions.target = "hiddenPage"	
		FrmPermissions.submit
		Refresh
	Else
		MsgBox "Please select an access permission to remove.", 0, "FNSNet Designer"		
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
		
	FrmPermissions.action = "UserCopy.asp"
	FrmPermissions.method = "GET"
	FrmPermissions.target = "hiddenPage"	
	FrmPermissions.submit
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

<SCRIPT LANGUAGE="JavaScript" FOR="VehicleBtnControl" EVENT="onscriptletevent (event, obj)">
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


<!--#include file="..\lib\UserBtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Locations</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmPermissions" ID="Form1">
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

IF UID <> "" Then  
       Set Conn = Server.CreateObject("ADODB.Connection")
	    Conn.Open CONNECT_STRING
	    SQLST1 = "SELECT ACCNT_HRCY_STEP_ID FROM ACCOUNT_USER WHERE USER_ID = " & UID
	       Set RS1 = Conn.Execute(SQLST1)
		          If Not RS1.EOF Then
			         RSASHID =  "" & RS1("ACCNT_HRCY_STEP_ID")
	              End If

           RS1.Close
		   Set RS1 = Nothing
    END IF 
%>

 <%If UID <> "" AND RSASHID <> "1"  Then
	  If UID <> "NEW" Then
		  Set Conn = Server.CreateObject("ADODB.Connection")
		  Conn.Open CONNECT_STRING
		  SQLST = "SELECT * FROM USERS_SITE_VIEW WHERE USER_ID = " & UID
		  Set RS = Conn.Execute(SQLST)
		  If Not RS.EOF Then
			           RSNAME =  "" & RS("NAME")
			   RSCREATIONDATE = "" & RS("PASSWORD_CREATION_DATE")
			 RSEXPIRATIONDATE = "" & RS("PASSWORD_EXPIRATION_DATE")
			RSLASTCHANGEDDATE = "" & RS("LAST_CHANGE_DATE")
		
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
<td><b>User ID:&nbsp;<span id="spanUID"><%=Request.QueryString("UID")%></span></td>
<td><b>Name:</b>&nbsp;<b><span id="spanName"><%=RSNAME%></span></b></td>
<td>&nbsp;&nbsp;</td>
<td>&nbsp;&nbsp;</td>
<td><b>Created Date:&nbsp;&nbsp;<span id="spanCreted"><%=RSCREATIONDATE%></span></td>
<td><b>Expiration Date:&nbsp;&nbsp;<span id="spanExpiration"><%=RSEXPIRATIONDATE%></span></td>
<td><b>Last Changed Date:&nbsp&nbsp;<span id="spanLastChanged"><%=RSLASTCHANGEDDATE%></span></td>
<td>&nbsp;&nbsp;</td>


</tr>
<tr>
</table>
<span class="Label" ID=SPANDATA>Retrieving...</span>


<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?NEWCAPTION=Add&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDEATTACH=TRUE&HIDEREMOVE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=VehicleBtnControl type=text/x-scriptlet VIEWASTEXT></OBJECT>
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="UserLocationsData.asp?<%=Request.QueryString%>">
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No valid account selected.
</div>


<% End If %>

</form>
</body>
</html>


