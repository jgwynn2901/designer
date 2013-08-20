<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
Response.Expires=0 
	AccountTextLen = 30	

	Dim SharedCount, SharedCountText, IBID
	SharedCount = 0
	SharedCountText = "Ready"
	
	IBID	= CStr(Request.QueryString("IBID"))

	If IBID <> "" Then
		If IBID = "NEW" Then 
		
			SharedCount = 0
		End If
	End If	
	
If IBID  <> "" Then
	If IBID  <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM INBOUNDCALL WHERE INBOUNDCALL_ID =" & IBID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RS_IBID        = RS("INBOUNDCALL_ID")
			RS_AHSID       = RS("ACCNT_HRCY_STEP_ID")
			RS_NAME        = RS("NAME")
			RS_PHONENUMBER = RS("PHONENUMBER")
			RS_DESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
			RS_GREETING    = ReplaceQuotesInText(RS("GREETING"))
			RS_ACCNT_TYPE = ReplaceQuotesInText(RS("ACCOUNT_TYPE_CD"))
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
End If
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Greeting Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT>
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();
var RuleSearchObj = new CRuleSearchObj();
var g_StatusInfoAvailable = false;
</SCRIPT>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if IBID <> "" then %>
			<% if SharedCount <= 1 then %>
			
		      document.all.txtPhoneNumber.value = "<%=RS_PHONENUMBER %>"
			  document.all.txtName.value = "<%=RS_NAME %>"
			  document.all.txtDescription.value = "<%=RS_DESCRIPTION%>"
			  document.all.txtGreeting.value = "<%=RS_GREETING%>"
			  document.all("txtAccType").value = "<%=RS_ACCNT_TYPE%>"
<%	else %>
	<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>

End Sub

Sub PostTo(strURL)
	FrmDetails.action = "GreetingSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateIBID(inIBID)
	document.all.IBID.value = inIBID
	document.all.spanIBID.innerText = inIBID
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

Function GetIBID
	if document.all.IBID.value <> "NEW" then
		GetIBID = document.all.IBID.value
	else
		GetIBID = ""
	end if 
End Function

'Function GetIBIDName
	'GetIBIDName = document.all.TxtName.value
'End Function

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
	
	
	If  document.all.txtName.value = "" then
		errmsg = errmsg &  "Name is a required field."
	end if
	If  document.all.AHSID_ID.innerText = "" then
		errmsg = errmsg &  "AHS ID is a required field."
	end if
	If  document.all.txtPhoneNumber.value = "" then
		errmsg = errmsg &  "PhoneNumber is a required field."
	end if
	If  document.all.txtGreeting.value = "" then
		errmsg = errmsg &  "Greeting is a required field."
	end if
	If errmsg = "" Then
		ValidateScreenData = true
	Else
		msgbox errmsg , 0 , "FNSDesigner"
		ValidateScreenData = false
	End If
End Function

sub UpdateScreenOnDelete()
	document.all.IBID.value = ""
	FrmDetails.action = "GreetingDetails.asp?STATUS=Delete successful."
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
	
	if document.all.IBID.value = "" then
		ExeDelete = false
		exit function
	end if

	if document.all.spanIBID.innerText = "" Then
		MsgBox "Inboundcall ID is " & document.all.spanIBID.innerText & vbCRLF & _
		"You cannot delete an attribute with no Inboundcall ID.",vbExclamation,"FNSNetDesigner"
		ExeDelete = false
		exit Function
	end if

	lret = Confirm("Are you sure you want to delete Inboundcall ID: " & document.all.IBID.value & " ?")

	if lRet = true Then
		document.all.TxtAction.value = "DELETE"
		sResult = document.all.IBID.value
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
	
	if document.all.IBID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.IBID.value = "NEW"
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
	
	if document.all.IBID.value = "" then
		ExeSave = false
		exit function
	end if
	
	
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
	
		If document.all.IBID.value = "NEW" then
		
			document.all.TxtAction.value = "INSERT"
		else
	
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if 
		
		sResult = sResult & "INBOUNDCALL_ID"& Chr(129) & document.all.IBID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.txtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONENUMBER"& Chr(129) & document.all.txtPhoneNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.txtDescription.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "GREETING"& Chr(129) & document.all.txtGreeting.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCOUNT_TYPE_CD"& Chr(129) & document.all.txtAccType.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	'Else
		'SpanStatus.innerHTML = "Nothing to Save"
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
Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Function AttachRule (ID, SPANID, strTITLE)
	RID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
			
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ESCALATION&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID, RuleSearchObj.RIDText
	ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID, RuleSearchObj.RIDText
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.value = ""
		SPANID.innerText = ""
	end if
End Function

Function DetachAccount(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Function AttachAccount (ID, SPANID)
	AHSID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID.title
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ESCALATION&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,AHSSearchObj ,"center"

	'if Selected=true update everything, otherwise if AHSID is the same, update text in case of save
	If AHSSearchObj.Selected = true Then
		If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = AHSSearchObj.AHSID
		end if
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	End If
End Function

<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Greeting Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="GreetingSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>

<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchPhoneNumber" value="<%=Request.QueryString("SearchPhoneNumber")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchGreeting" value="<%=Request.QueryString("SearchGreeting")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="IBID" value="<%=Request.QueryString("IBID")%>" >

<%	

If IBID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</td>
</tr>
</table>
<table class="LABEL">
<tr>
	<td>
	<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::DetachAccount AHSID_ID, AHSID_TEXT">
	</td>
	<td width=305 nowrap>Account:<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(RSACCOUNT_NAME)%>" ><%=TruncateText(RSACCOUNT_NAME,AccountTextLen)%></SPAN></td>
	<td>A.H.Step ID:<SPAN ID=AHSID_ID CLASS=LABEL><%=RS_AHSID%></SPAN></td>
	</tr>
</table>

<table CLASS="LABEL" CELLPADDING=0 CELLSPACING=0 >
<tr>
<td>
<td><table class="LABEL">
<TR>
	<td CLASS=LABEL COLSPAN=2>Inbound Call ID:&nbsp<span id="spanIBID"><%=Request.QueryString("IBID")%></span></td>
	</tr>
	</TABLE>
	<TABLE>
	<tr>
	<tr nowrap>
    <td CLASS=LABEL >Name:<br>
	<input size=45 
	       CLASS="LABEL"
	       MAXLENGTH=80 
	       TYPE="TEXT" 
	       NAME="txtName" 
	       tabindex=1
	       VALUE="<%=RS_NAME %>" 
	       ONKEYPRESS="VBScript::Control_OnChange"
	       ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS=LABEL >Phone Number:<br>
	<input  size=30
	       CLASS="LABEL"
	       MAXLENGTH=40
	       TYPE="TEXT" 
	       NAME="txtPhoneNumber"
	       tabindex=2 
	       VALUE="<%=RS_PHONENUMBER %>" 
	       ONKEYPRESS="VBScript::Control_OnChange" 
	       ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
    <tr>
	<td CLASS="LABEL" colspan=3>Description:<br>
	<input size="80"
	      CLASS="LABEL" 
	       MAXLENGTH=255 
	       TYPE="TEXT" 
	       NAME="txtDescription"
	       tabindex=3
	       VALUE="<%=RS_DESCRIPTION%>" 
	       ONKEYPRESS="VBScript::Control_OnChange"
	       ONCHANGE="VBScript::Control_OnChange"></td>
	
	<tr>
	<td CLASS="LABEL"  colspan=3>Greeting:<br>
    <input size=80
	       CLASS="LABEL" 
	       MAXLENGTH=2000 
	       TYPE="TEXT" 
	       NAME="txtGreeting" 
	       tabindex=4
	       VALUE="<%=RS_GREETING%>"
	      ONKEYPRESS="VBScript::Control_OnChange" 
	      ONCHANGE="VBScript::Control_OnChange"></td>
    </tr> 
	<tr>
	<td CLASS="LABEL"  colspan=3>Account Type Code:<br>
	<select STYLE="WIDTH:75" NAME="txtAccType" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Select1">
		<option VALUE="I">I
		<option VALUE="T">TPA
	</select>
    </tr> 
 </table>
<% Else %>
<DIV style="margin-top:170px;margin-left:170px" CLASS="LABEL">
<%=Request.QueryString("STATUS") & "<br>"%>
No Greeting selected.
</DIV>
<% End If %>
</form>
</body>
</html>


