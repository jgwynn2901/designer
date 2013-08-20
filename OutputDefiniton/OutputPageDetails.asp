<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	Response.Expires=0 
	If Request.QueryString("ODID") <> "" THEN RSOUTPUTDEF_ID = Request.QueryString("ODID")
	Dim SharedCount, SharedCountText, OPID
	SharedCount = 0
	SharedCountText = "Ready"
	OPID	= CStr(Request.QueryString("OPID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Output Page Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function COutputDefinitionSearchObj()
{
	this.ODID = "";
	this.ODIDName = "";
	this.Saved = false;	
	this.Selected = false;	
}

function LaunchODEditor(key) {
Url = "../RoutingPlan/OutputDefinitionEditor-f.asp?AHSID=<%= Request.QueryString("AHSID") %>&" + key
var VisEditorObj = window.open(Url, null, "height=500,width=750,status=no,toolbar=no,menubar=no,location=no,resizable=yes,top=0,left=0");
lret = Handles( VisEditorObj, "OUTPUT");
VisEditorObj.focus()
}
var DefinitionObj = new COutputDefinitionSearchObj();
var g_StatusInfoAvailable = false;
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Function Handles(Obj, Title)
	If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
		lret = top.frames("TOP").SetHandle(Obj, Title)
	End If
End Function

Function ExeLaunch
<% If Request.QueryString("OPID") <> "" Then %>
key = ""
key = key & "ODID=" & document.all.TxtOUTPUTDEF_ID.value
key = key & "&OPID=" & trim(document.all.spanOPID.innerhtml)

If document.all.spanOPID.innerhtml <> "NEW" Then
	Call LaunchODEditor(key)
Else
	msgbox "Please save this Output Page before launching the visual editor." , 0, "FNSDesigner"
End If
<% End if %>
End Function

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if OPID <> "" then
			if SharedCount <= 1 then %>
	document.all.ChkEdit.checked = true
	ChkEdit_OnClick

<%			else %>
	document.all.ChkEdit.checked = false
	ChkEdit_OnClick
	SetStatusInfoAvailableFlag(true)
<%
				SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
						
		end if	'OPID <> ""

	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "OutputPageSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateOPID(inOPID)
	document.all.OPID.value = inOPID
	document.all.spanOPID.innerText = inOPID
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

Function GetOPID
	if document.all.OPID.value <> "NEW" then
		GetOPID = document.all.OPID.value
	else
		GetOPID = ""
	end if 
End Function

Function GetOPIDName
	GetOPIDName = document.all.TxtName.value
End Function

Function GetOPIDCaption
	GetOPIDCaption = document.all.TxtCaption.value
End Function

Function GetOPIDInputType
	GetOPIDInputType = document.all.TxtInputType.value
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
errmsg = ""
	If  document.all.TxtName.value = "" then
		errmsg = errmsg & "Name is a required field." & VbCrLf
	end if
	If  document.all.TxtOUTPUTDEF_ID.value = "" then
		errmsg = errmsg & "Output Definition ID is a required field."& VbCrLf
	end if
	if document.all.TxtPAGE_NUMBER.value = "" OR Not IsNumeric(document.all.TxtPAGE_NUMBER.value)  then
			errmsg = errmsg & "Page Number is required and must be numeric."& VbCrLf
	end if
	If errmsg = "" Then
		ValidateScreenData = true
	Else
		msgbox errmsg, 0, "FNSDesigner"
	End If
End Function

sub UpdateScreenOnDelete()
	document.all.OPID.value = ""
	FrmDetails.action = "OutputPageDetails.asp?STATUS=Delete successful."
	FrmDetails.target = "_self"
	FrmDetails.submit
end sub

Function ExeDelete
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeDelete = bRet
		exit Function
	end if
	
	if document.all.OPID.value = "" then
		ExeDelete = false
		exit function
	end if

	lret = Confirm("Are you sure you want to delete Output Page ID: " & document.all.OPID.value & " ?")

	if lRet = true Then
		document.all.TxtAction.value = "DELETE"
		sResult = document.all.OPID.value
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		ExeDelete = true
	Else
		ExeDelete = false
	End if
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.OPID.value = "" then
		ExeCopy = false
		exit function
	end if
	document.all.OPID.value = "NEW"
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
	
	if document.all.OPID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.OPID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		'document.all.TxtName.value = Replace(document.all.TxtName.value," ","")		
		'document.all.TxtName.value = UCase(document.all.TxtName.value)		
		sResult = sResult & "OUTPUT_PAGE_ID"& Chr(129) & document.all.OPID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OUTPUTDEF_ID"& Chr(129) & document.all.TxtOUTPUTDEF_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtNAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PAGE_NUMBER"& Chr(129) & document.all.TxtPAGE_NUMBER.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OUTPUT_TRAY"& Chr(129) & document.all.TxtOUTPUT_TRAY.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "BACKGROUND_BMP"& Chr(129) & document.all.TxtBACKGROUND_BMP.value & Chr(129) & "1" & Chr(128)
		if document.all.TxtOrientation.checked then 
		   sResult = sResult & "ORIENTATION"& Chr(129) & "L" & Chr(129) & "1" & Chr(128)
		else
		   sResult = sResult & "ORIENTATION"& Chr(129) & "P" & Chr(129) & "1" & Chr(128)
		end if
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

Sub BtnFindOD_onclick
SetDirty()
	lret = window.showModalDialog( "../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=MODAL"  ,DefinitionObj ,"dialogWidth:450px;dialogHeight:450px;center")
	if DefinitionObj.ODID <> "" Then
		document.all.TxtOUTPUTDEF_ID.value = DefinitionObj.ODID
	end if
End Sub

Sub VEditor_OnClick

key = ""
key = key & "ODID=" & document.all.OUTPUTDEF_ID.value
key = key & "&OPID=<%= OUTPUT_PAGE_ID %>" 
Call LaunchODEditor(key)
End Sub

Function Handles(Obj, Title)
If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
	lret = top.frames("TOP").SetHandle(Obj, Title)
End If
End Function
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Output Page Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="OutputPageSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchOPID" value="<%=Request.QueryString("SearchOPID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchODID" value="<%=Request.QueryString("SearchODID")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" name="SearchBMP" value="<%=Request.QueryString("SearchBMP")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="OPID" value="<%=Request.QueryString("OPID")%>">

<%	

Function TruncateRuleText(inText)
	if not IsNull(inText) then
		If Len(inText) < 40 Then
			TruncateRuleText = inText
		Else
			TruncateRuleText = Mid ( inText, 1, 40) & " ..."
		End If
	end if
End Function

Function TruncateLookupText(inText)
	if not IsNull(inText) then
		If Len(inText) < 22 Then
			TruncateLookupText = inText
		Else
			TruncateLookupText = Mid ( inText, 1, 22) & " ..."
		End If
	end if
End Function

Function ReplaceRuleText(inText)
	if not IsNull(inText) then
		ReplaceRuleText = Replace(inText,"""","&quot;")
	end if
End Function

If OPID <> "" Then
	If OPID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM OUTPUT_PAGE WHERE OUTPUT_PAGE_ID = " & OPID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSOUTPUT_PAGE_ID = RS("OUTPUT_PAGE_ID")
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSOUTPUTDEF_ID = RS("OUTPUTDEF_ID")
			RSPAGE_NUMBER = RS("PAGE_NUMBER")
			RSOUTPUT_TRAY = ReplaceQuotesInText(RS("OUTPUT_TRAY"))
			RSBACKGROUND_BMP = RS("BACKGROUND_BMP")
			RSORIENTATION = RS("ORIENTATION")
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
	
%>
<table ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
<input ScrnBtn="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>
	<table class="LABEL">
	<tr>
	<td>Output Page ID:&nbsp;<span id="spanOPID"><%=Request.QueryString("OPID")%></span></td>
	</tr>
	<tr>
	<td CLASS="LABEL" VALIGN="BOTTOM">Output Definition:<br><input READONLY ScrnInput="TRUE" TYPE="TEXT" SIZE="10" CLASS="LABEL" NAME="TxtOUTPUTDEF_ID" VALUE="<%= RSOUTPUTDEF_ID %>" STYLE="BACKGROUND-COLOR:SILVER" VALUE>
	<img SRC="../IMAGES/Attach.gif" ID="BtnFindOD" TITLE="Assign Output Definition" STYLE="CURSOR:HAND" align="absbottom" WIDTH="16" HEIGHT="16"></td>
	</tr>
	<tr>
	<td>Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="255" SIZE="60" TYPE="TEXT" NAME="TxtName" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Page Number:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtPAGE_NUMBER" VALUE="<%=RSPAGE_NUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td>Background Bmp:<br><input ScrnInput="TRUE" size="60" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtBACKGROUND_BMP" VALUE="<%=RSBACKGROUND_BMP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Output Tray:<br><input ScrnInput="TRUE" size="30" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtOUTPUT_TRAY" VALUE="<%=RSOUTPUT_TRAY%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<%if RSORIENTATION = "L" then %>
	   <td><input ScrnBtn="False" ScrnInput="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="TxtOrientation" ID="Checkbox1"  checked  value="L"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" >Landscape</td>
	<%else%>
	    <td><input ScrnBtn="False" ScrnInput="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="TxtOrientation" ID="Checkbox2"   value="P"   ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" >Landscape</td>
	<%end if%>
	</tr>
	
	</table>
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
<%=Request.QueryString("STATUS") & "<br>"%>
No Output Page selected.
</div>


<% End If %>

</form>
</body>
</html>


