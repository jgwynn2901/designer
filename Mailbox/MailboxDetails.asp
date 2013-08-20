<%
'***************************************************************
'form for Mailbox data entry.
'
'$History: MailboxDetails.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:45p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Mailbox
'* Hartford SRS: Initial revision
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	Response.Expires=0 %>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Branch Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
var g_StatusInfoAvailable = false;

function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}
</script>
<script language =vbscript >
Sub window_onload
dim cInnerHTML

<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
	
<%	end if %>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "MailboxSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateMBID(inMBID)
	document.all.MBID.value = inMBID
	document.all.spanMBID.innerText = inMBID
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

Function GetMBID
	if document.all.MBID.value <> "NEW" then
		GetMBID = document.all.MBID.value
	else
		GetMBID = ""
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
	If  document.all.TxtAHLoadID.value = "" then
		MsgBox "AH Load ID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	if IsNumeric(document.all.TxtAHLoadID.value) = false then
		MsgBox "Please enter a number in the AH Load ID field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.TxtMailboxNumber.value = "" then
		MsgBox "Mailbox Number is a required field.",0,"FNSNetDesigner"
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
	
	if document.all.MBID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.MBID.value = "NEW"
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
	
	if document.all.MBID.value = "" then
		ExeSave = false
		exit function
	end if
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.MBID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "MAILBOX_ID"& Chr(129) & document.all.spanMBID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAILBOX_NUMBER"& Chr(129) & document.all.TxtMailboxNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCOUNT_HIERARCHY_LOAD_ID"& Chr(129) & document.all.TxtAHLoadID.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
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

sub SetBranchTypeFieldReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("SpecialFilterBtn") = "TRUE" then
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Mailbox Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="MailboxSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchMBID" value="<%=Request.QueryString("SearchMBID")%>">
<input type="hidden" name="SearchMailboxNumber" value="<%=Request.QueryString("SearchMailboxNumber")%>">
<input type="hidden" name="SearchAHLoadID" value="<%=Request.QueryString("SearchAHLoadID")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="MBID" value="<%=Request.QueryString("MBID")%>">
<%	

MBID = CStr(Request.QueryString("MBID"))
If MBID <> "" Then
	If MBID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM MAILBOX WHERE MAILBOX_ID = " & MBID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then 
			RSMAILBOX_NUMBER = ReplaceQuotesInText(RS("MAILBOX_NUMBER"))
			RSACCOUNT_HIERARCHY_LOAD_ID = RS("ACCOUNT_HIERARCHY_LOAD_ID")
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
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

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table class="LABEL">
	<tr>
	<tr>
	<tr>
	<tr>
	<td COLSPAN="4">Mailbox ID:&nbsp;<span id="spanMBID"><%=Request.QueryString("MBID")%></span></td>
	<td> 
	</tr> 
	<tr>
	<td COLSPAN="2">Mailbox Number:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtMailboxNumber" VALUE="<%=RSMAILBOX_NUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>AH Load ID:<br><input ScrnInput="TRUE" size="12" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtAHLoadID" VALUE="<%=RSACCOUNT_HIERARCHY_LOAD_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	</table>
</td>
</tr> 
</table>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Mailbox selected.
</div>

<% End If %>

</form>
</body>
</html>


