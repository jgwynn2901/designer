<%
'***************************************************************
'form for Mailbox data entry.
'
'$History: MyGreetingDetails.asp $ 
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:35p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:31p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:27p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* took out stop
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:25p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:14p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:09p
'* Created in $/FNS_DESIGNER/Source/Designer/Greeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 4/21/08    Time: 9:23a
'* Created in $/FNS_DESIGNER/Source/Designer
'* created for Sedgwick.  Just want to save my work for now
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
<title>Greeting Details</title>
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
	FrmDetails.action = "MyGreetingSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateGreetingID(inMBID)
	document.all.GreetingID.value = inMBID
	document.all.spanGreetingID.innerText = inMBID
End Sub

Sub UpdateGreetingText(inMBID)
	document.all.TxtGreetingText.value = inMBID
	document.all.SpanGreetingText.innerText = inMBID
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

Function GetGreetingID
	if document.all.GreetingID.value <> "NEW" then
		GetGreetingID = document.all.GreetingID.value
	else
		GetGreetingID = ""
	end if 
End Function

Function GetGreetingText
	if document.all.TxtGreetingText.value <> "NEW" then
		GetGreetingText = document.all.TxtGreetingText.value
	else
		GetGreetingText = ""
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

	If  document.all.TxtCONTRACTNUMBER.value = "" then
		MsgBox "Contract Number is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	if document.all.TxtEmployeeFeedFlag.value = "" then
		MsgBox "Employee Feed Flag is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.TxtGreetingText.value = "" then
		MsgBox "Greeting Text is a required field.",0,"FNSNetDesigner"
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
	
	if document.all.GreetingID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.GreetingID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeDelete
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeDelete = bRet
		exit Function
	end if
	
	if document.all.GreetingID.value = "" then
		ExeDelete = false
		exit function
	end if

	''if lRet = true Then
		document.all.TxtAction.value = "DELETE"
		sResult = document.all.AID.value
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		ExeDelete = true
	''Else
	''	ExeDelete = false
	'End if
End Function


Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.GreetingID.value = "" then
		ExeSave = false
		exit function
	end if
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.GreetingID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		
	
		sResult = sResult & "CONTRACT_NUM"& Chr(129) & document.all.TxtContractNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TEXT"& Chr(129) & document.all.TxtGreetingText.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CODES"& Chr(129) & document.all.TxtLob.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "HAS_EMPLOYEE_FEED"& Chr(129) & document.all.TxtEmployeeFeedFlag.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "GREETINGS_ID"& Chr(129) & document.all.GreetingId.value & Chr(129) & "1" & Chr(128)
		
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Greeting Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="MyGreetingSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchGreetingID" value="<%=Request.QueryString("SearchGreetingID")%>">
<!--<input type="hidden" name="SearchMailboxNumber" value="<%=Request.QueryString("SearchMailboxNumber")%>">-->
<input type="hidden" name="SearchContractNumber" value="<%=Request.QueryString("SearchContractNumber")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="GreetingID" value="<%=Request.QueryString("GreetingID")%>">
<%	

GreetingID = CStr(Request.QueryString("GreetingID"))
If GreetingID <> "" Then
	If GreetingID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM GREETINGS WHERE GREETINGS_ID = " & GreetingID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then 
			RSGREETINGS_ID = RS("GREETINGS_ID")
			RSCONTRACT_NUMBER = RS("CONTRACT_NUM")
			RSTEXT = RS("TEXT")
			RSFEED = RS("HAS_EMPLOYEE_FEED")
			RSLOB = RS("LOB_CODES")
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

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0" >
<tr>
<td> 
	<table class="LABEL" >
	<tr>
	<tr>
	<tr>
	<tr>
	<td COLSPAN="7">Greeting ID:&nbsp;<span id="spanGreetingID"><%=Request.QueryString("GreetingID")%></span></td>
	<td> 
	</tr> 
	<tr>
	<td COLSPAN="7">Greeting Text:<br><span id="SpanGreetingText"><Textarea  NAME="TxtGreetingText" cols = "70" rows = "15"   ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"><%=RSTEXT%></textarea></td>
	</tr>
	<tr>
		<td colspan = "7">LOB:<br><input ScrnInput="TRUE" size="120" CLASS="LABEL" MAXLENGTH="120" TYPE="TEXT" NAME="TxtLob" VALUE="<%=RSLOB%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text1"></td>
		<td>Employee Feed:<br><SELECT  SIZE = "1"  ScrnBtn="TRUE"  CLASS="LABEL"  NAME="TxtEmployeeFeedFlag"  ONCHANGE="VBScript::Control_OnChange">
		<option value = " " selected> </option>
		<option value = "Y">Y</option>
		<option value = "N">N</option>
		
		</select>
		</td>
					
	</tr>
<tr>	
	<td >Contract Number:<br><input ScrnInput="TRUE" size="4" CLASS="LABEL" MAXLENGTH="4" TYPE="TEXT" NAME="TxtCONTRACTNUMBER" VALUE="<%=RSCONTRACT_NUMBER%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text2">
	</td>
	</tr>
	</table>
</td>
</tr> 
</table> 

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Greetings selected.
</div>

<% End If %>

<%If Not IsNull(RSFEED) Then
	If  CStr(RSFEED) <> "" Then	 %>
<SCRIPT LANGUAGE="VBScript">
	SelectOption document.all.TxtEmployeeFeedFlag,"<%=CStr(RSFEED)%>"
</SCRIPT>
<%	End If
End If  %>

</form>
</body>
</html>


