<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	RID = Request.QueryString("RID")
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Dictionary</title>

<script LANGUAGE="VBScript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false


sub window_onload
<%	
	if CStr(Request.QueryString("MODE")) = "RO" then %>
	   SetScrnInputsReadOnly true,"DISABLED"
	   SetScrnBtnsReadOnly true
<%	else
        
		if RID <> "" then %>
		        document.all.ChkEdit.checked = true
				ChkEdit_OnClick
				SetStatusInfoAvailableFlag(true)	
		<%	
		end if	'RID <> ""
	end if %>
End Sub

Sub PostTo(strURL)
	FrmDictionaryEditor.action = "DictionarySearch-f.asp"
	FrmDictionaryEditor.method = "GET"
	FrmDictionaryEditor.target = "_parent"	
	FrmDictionaryEditor.submit
End Sub

sub SetRID(inRID)
	document.all.RID.value = inRID
end sub

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


function GetDictText
	GetDictText = document.all.TxtDictText.value
end function

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
strError = ""
	If document.all.TxtDictText.value = "" Then	strError = "Word is a required field." & VBCRLF
	
	If strError <> "" Then
		ValidateScreenData = false
		MsgBox strError, 0 , "FNSNetDesigner"
	Else
		ValidateScreenData = true
	End If
End Function

Function ExeSave
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = false
		Exit Function
	End If
	
	If document.all.RID.value = "" Then
		ExeSave = false
		Exit Function
	End If
	
'	If CStr(document.body.getAttribute ("ScreenDirty")) = "YES" Then
		If ValidateScreenData = false Then 
			ExeSave = false
			Exit Function
		End If
		
		FrmDictionaryEditor.submit		
		ExeSave = true
'	Else
'		document.all.SpanStatus.innerHTML = "Nothing to Save" 		
'		ExeSave = false
'	End If
End Function

Function ExeCopy
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.RID.value = "" Then
		ExeCopy = false
		Exit Function
	End If

	document.body.setAttribute "ScreenDirty","YES"
	document.all.RID.value = "NEW"
	
	ExeCopy = ExeSave
End Function

Function ExeDelete
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = false
		Exit Function
	End If
	
	If document.all.RID.value = "" Then
		ExeDelete = false
		Exit Function
	End If
    document.all.DeleteFlag.value = "Y"
	FrmDictionaryEditor.submit		
	ExeDelete = true
		
End Function

sub SetScrnInputsReadOnly(bReadOnly, strNewClass)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnInput") = "TRUE" then
			document.all(iCount).readOnly = bReadOnly
			document.all(iCount).className = strNewClass
		end if
	next
end sub

sub SetScrnBtnsReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next
end sub

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"

	if document.all.ChkEdit.checked = true then
		SetScrnInputsReadOnly false,"LABEL"
		SetScrnBtnsReadOnly false
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScrnInputsReadOnly true,"DISABLED"
		SetScrnBtnsReadOnly true
		document.body.setAttribute "ScreenMode", "RO"				
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"	
end sub

sub Control_OnChange
<%	if CStr(Request.QueryString("MODE")) <> "RO" then %>
	document.body.setAttribute "ScreenDirty", "YES"	
<%	end if %>
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
<form name="FrmDictionaryEditor" method="POST" action="../dictionary/dictionaryEditorExe.asp" target="hiddenPage">
<% 'need to maintain these values in order to post back to the search tab %>

<input type="hidden" name="SearchDictText" value="<%=Request.QueryString("SearchDictText")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="RID" value="<%=Request.QueryString("RID")%>">
<input type="hidden" name="DeleteFlag" value=false>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Dictionary Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>


<%	If RID <> "" then
		If RID <> "NEW" Then
			strExecute = "SELECT * FROM SPELL_CHECK WHERE WORD = '" & RID & "'"
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			Set rs = Conn.Execute(strExecute)

			if not rs.EOF then
				DictText = rs("WORD")
			end if
			
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing		
		End If 'RID <> "NEW"

%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
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

<table CLASS="LABEL" ALIGN="CENTER" width="100%">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td>Word:</td></tr>
<tr>
<!--<td input class="LABEL" name="SearchDictText" type="text" size="16" VALUE="<%=Request.QueryString("RID")%>">
</TD>-->

<td><input type="text" name="txtDicttext" ScrnInput="TRUE" CLASS="LABEL" VALUE="<%=Request.QueryString("RID")%>"></td>
</tr>
</table>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Word selected.
</div>


<% End If %>


</body>
</form>
</html>

