<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\CheckSharedLUType.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	Response.Expires=0 

	Dim SharedCount, SharedCountText, LUTID
	SharedCount = 0
	SharedCountText = "Ready"
	
	LUTID	= CStr(Request.QueryString("LUTID"))
	
	If LUTID <> "" Then
		If LUTID = "NEW" Then 
			SharedCount = 0
		Else
			SharedCount = CheckSharedLUType(CLng(LUTID),True,True,1,False,False,0)
		End If
	End If	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Lookup Type Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CLookupCodeSearchObj()
{
	this.Selected = false;
}
var LookupCodeSearchObj = new CLookupCodeSearchObj();
var g_StatusInfoAvailable = false;

</script>
<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	else 
		if LUTID <> "" then
			if SharedCount <= 1 then %>
				document.all.ChkEdit.checked = true;
				ChkEdit_OnClick();
			<%else %>
				document.all.ChkEdit.checked = false;
				ChkEdit_OnClick();
				SetStatusInfoAvailableFlag(true);
				<%SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
				 if CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
					document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" + "<Font size=1 Color='Maroon'>+</Font>";
				<%else %>
					document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>";
				<%end if
			end if		
		end if	'LUTID <> ""
	end if 
%>
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 150;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 150;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmDetails.action = "LookupTypeSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateLUTID(inLUTID)
	document.all.LUTID.value = inLUTID
	document.all.spanLUTID.innerText = inLUTID
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

Function GetLUTID
	if document.all.LUTID.value <> "NEW" then
		GetLUTID = document.all.LUTID.value
	else
		GetLUTID = ""
	end if 
End Function

Function GetLUTIDName
	GetLUTIDName = document.all.TxtName.value
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
	If  document.all.TxtName.value = "" then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	ValidateScreenData = true
End Function

Function GetSelectedLUCID
	GetSelectedLUCID = document.frames("DataFrame").GetSelectedLUCID
End Function

Sub ExeNewCode
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.LUTID.value = "" Or document.all.LUTID.value = "NEW" Then
		Exit Sub
	End If

<%If HasAddPrivilege("FNSD_LOOKUP_TYPES","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to add lookup codes.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		


	dim LUTID, LUCID, MODE
	LUCID = "NEW"
	LUTID = document.all.LUTID.value
	MODE = document.body.getAttribute("ScreenMode")
	
	LookupCodeSearchObj.Selected = false

	strURL = "LookupCodeModal.asp?LUTID=" & LUTID & "&LUCID=" & LUCID & "&MODE=" & MODE 	
	showModalDialog  strURL,LookupCodeSearchObj ,"center"

	If LookupCodeSearchObj.Selected = true Then	Refresh
End Sub

Sub Refresh
	LUTID = document.all.LUTID.value
	document.all.tags("IFRAME").item("DataFrame").src = "LookupTypeDetailsData.asp?LUTID=" & LUTID
End Sub

Sub ExeEditCode
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.LUTID.value = "" Or document.all.LUTID.value = "NEW" Then
		Exit Sub
	End If
	
	dim LUCID, LUTID
	LUCID = GetSelectedLUCID
	LUTID = document.all.LUTID.value
	
	If LUCID <> "" Then
		LookupCodeSearchObj.Selected = false
		strURL = "LookupCodeModal.asp?LUTID=" & LUTID & "&LUCID=" & LUCID & "&MODE=" & MODE 	
		showModalDialog  strURL,LookupCodeSearchObj ,"center"
		If LookupCodeSearchObj.Selected = true Then	Refresh
	Else
		MsgBox "Please select a lookup code to Edit.", 0, "FNSNet Designer"		
	End If
	
End Sub

Sub ExeRemoveCode
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.LUTID.value = "" Or document.all.LUTID.value = "NEW" Then
		Exit Sub
	End If

<%If HasDeletePrivilege("FNSD_LOOKUP_TYPES","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to delete lookup codes.",0,"FNSNetDesigner"
		Exit Sub
<%End If %>		

	dim LUCID, sResult
	LUCID = GetSelectedLUCID
	
	If LUCID <> "" Then
		sResult = sResult &  LUCID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmDetails.action = "LookupCodeSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a lookup code to Remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub

Function InEditMode
	InEditMode = true
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		InEditMode = false
	End If

End Function

Function ExeCopy
	If Not InEditMode Then
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.LUTID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
	
	document.all.SpanSharedCount.innerText = 0
	
	FrmDetails.action = "LookupTypeCopy.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit
'	Refresh is done inside LookupTypeCopy.asp due to timing
	ExeCopy = true
End Function

Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.LUTID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.LUTID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "LU_TYPE_ID"& Chr(129) & document.all.LUTID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "LookupTypeSave.asp"
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

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
	End If		
End Sub

Sub RefCountRpt_onclick()
	If document.all.SpanSharedCount.innerText > 0 Then
		If document.all.LUTID.value <> "" And document.all.LUTID.value <> "NEW" Then
			paramID = document.all.LUTID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedLUType=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
	Else
		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
	End If	
End	Sub


Sub RefCountRpt_onmouseover()
	If document.all.SpanSharedCount.innerText > 0 Then
		document.all.RefCountRpt.style.cursor = "HAND"
	Else
		document.all.RefCountRpt.style.cursor = "DEFAULT"
	End If
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
<!--#include file="..\lib\LTBtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Lookup Type Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="LookupTypeSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchLUTID" value="<%=Request.QueryString("SearchLUTID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="LUTID" value="<%=Request.QueryString("LUTID")%>">

<%	

If LUTID <> "" Then
	If LUTID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM LU_TYPE WHERE LU_TYPE_ID = " & LUTID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSNAME = ReplaceQuotesInText(RS("NAME"))
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>
<table ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td WIDTH="14">
<img ID="RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count">
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">
:<span id="SpanSharedCount"><%=SharedCount%></span>
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="455">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
<input ScrnBtn="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>

<table CLASS="LABEL" >
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td>Lookup Type ID:&nbsp;<span id="spanLUTID"><%=Request.QueryString("LUTID")%></span></td></tr>
<tr><td>Name:<br><input ScrnInput="TRUE" MAXLENGTH="80" CLASS="LABEL" size="60" TYPE="TEXT" NAME="TxtName" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
</table>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Lookup Codes</td>
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
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&HIDEATTACH=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=LTBtnControl type=text/x-scriptlet></OBJECT>
<iframe  FRAMEBORDER="0" width=100% height=0 name="DataFrame" src="LookupTypeDetailsData.asp?<%=Request.QueryString%>">
</fieldset>
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No lookup type selected.
</div>


<% End If %>

</form>
</body>
</html>


