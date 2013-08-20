<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%Response.Expires=0 
	Dim SharedCount, SharedCountText, ABID
	SharedCount = 0
	SharedCountText = "Ready"
	
	ABID	= CStr(Request.QueryString("ABID"))
	If ABID <> "" Then
		If ABID = "NEW" Then 
			SharedCount = 0
		End If
	End If	
	

If ABID <> "" Then
	If ABID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ADDRESS_BOOK_ENTRY WHERE ADDRESS_BOOK_ENTRY_ID = " & ABID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSADDRESS_BOOK_ENTRY_ID = RS("ADDRESS_BOOK_ENTRY_ID")
			RSCALLFLOW_ID = RS("CALLFLOW_ID")
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
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
<title>Address Book Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function CRPSearchObj()
{
	this.routing_plan_id = "";
	this.ahsid = "";
	this.multiselected = "";
}
var SearchObj = new CRPSearchObj();
-->
</SCRIPT>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if ABID <> "" then %>
			<% if SharedCount <= 1 then %>
			
<%	else %>
	SetStatusInfoAvailableFlag(true)
			
<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "AddressBookSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"
	FrmDetails.submit
End Sub

Sub UpdateABID(inABID)
	document.all.ABID.value = inABID
	document.all.spanABID.innerText = inABID
End Sub

Function ExeDelete
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeDelete = bRet
		exit Function
	end if
	
	if document.all.ABID.value = "" then
		ExeDelete = false
		exit function
	end if

	lret = Confirm("Are you sure you want to delete Address Book ID: " & document.all.ABID.value & " ?")

	if lRet = true Then
		document.all.TxtAction.value = "DELETE"
		sResult = document.all.ABID.value
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		ExeDelete = true
	Else
		ExeDelete = false
	End if
End Function

sub UpdateScreenOnDelete()
	document.all.ABID.value = ""
	FrmDetails.action = "AddressBookDetails.asp?STATUS=Delete successful."
	FrmDetails.target = "_self"
	FrmDetails.submit
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

Function GetABID
	if document.all.ABID.value <> "NEW" then
		GetABID = document.all.ABID.value
	else
		GetABID = ""
	end if 
End Function

Function GetABIDName
	GetABIDName = document.all.TxtName.value
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
	If  document.all.TxtCALLFLOW_ID.value = "" then
		errmsg = errmsg & "Call Flow ID is a required field."
	end if
	if document.all.TxtName.value = "" Then
		errmsg = errmsg & "Name is a required field."
	end if
	if errmsg = "" Then
		ValidateScreenData = true
	Else
		msgbox errmsg, 0, "FNSDesigner"
		ValidateScreenData = false
	End IF
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.ABID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.ABID.value = "NEW"
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
	
	if document.all.ABID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.ABID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "ADDRESS_BOOK_ENTRY_ID"& Chr(129) & document.all.ABID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "CALLFLOW_ID"& Chr(129) & document.all.TxtCALLFLOW_ID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtNAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDESCRIPTION.value & Chr(129) & "1" & Chr(128)
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


Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

'Sub RefCountRpt_onclick()
'	If document.all.SpanSharedCount.innerText > 0 Then
'		If document.all.ABID.value <> "" And document.all.ABID.value <> "NEW" Then
'			paramID = document.all.ABID.value
'		Else	
'			paramID = 0
'		End If
'		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedAttribute=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
'	Else
'		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
'	End If	
'End	Sub
'Sub RefCountRpt_onmouseover()
'	If document.all.SpanSharedCount.innerText > 0 Then
'		document.all.RefCountRpt.style.cursor = "HAND"
'	Else
'		document.all.RefCountRpt.style.cursor = "DEFAULT"
'	End If
'End Sub

Sub BtnAttachVehicle_OnClick
	MODE = document.body.getAttribute("ScreenMode")
If MODE = "RW" Then
	VehicleObj.VehicleID = VehicleID
	strURL = "VehicleMaintenance.asp?CONTAINERTYPE=MODAL&SEARCHONLY=TRUE"
	showModalDialog  strURL  ,VehicleObj ,"dialogWidth=600px; dialogHeight=500px; center=yes"
	If VehicleObj.VehicleID <> "" Then
			document.body.setAttribute "ScreenDirty", "YES"	
			document.all.TxtVEHICLE_ID.value = VehicleObj.VehicleID
	End If
End If
End Sub

Sub BtnAttachCF_OnClick
If document.body.GetAttribute("ScreenMode") = "RO" Then
	Exit Sub
End If

	lret = ""
	strURL = ""
	SearchObj.multiselected = ""
	strURL = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=MODAL&LAUNCHER=SEARCH&SECURITYPRIV=FNSD_ADDRESS_BOOK"
	lret = window.showModalDialog(strURL  ,SearchObj ,"dialogWidth:550px;dialogHeight:550px;center")
	if SearchObj.multiselected <> "" Then
		document.all.TxtCALLFLOW_ID.value = SearchObj.multiselected
	End If	
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Address Book Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="AddressBookSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" name="SearchABID" value="<%=Request.QueryString("SearchABID")%>">
<input type="hidden" name="SearchNAME" value="<%=Request.QueryString("SearchNAME")%>">
<input type="hidden" name="SearchDESCRIPTION" value="<%=Request.QueryString("SearchDESCRIPTION")%>">
<input type="hidden" name="SearchCFID" value="<%=Request.QueryString("SearchCFID")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="ABID" value="<%=Request.QueryString("ABID")%>" >

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
If ABID <> "" Then

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
<table CLASS="LABEL" CELLPADDING=0 CELLSPACING=0 >
<tr><td>
<table class="LABEL">
	<tr>
	<td COLSPAN=5 CLASS=LABEL>Address Book ID:&nbsp<span id="spanABID"><%=Request.QueryString("ABID")%></span></td>
	</TR>
	<TR>
	<TD VALIGN=BOTTOM>Call Flow ID:<BR><IMG NAME=BtnAttachCF STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Call Flow">
	<input ScrnInput="TRUE" size=10 CLASS="LABEL" READONLY STYLE="BACKGROUND-COLOR:SILVER" MAXLENGTH=10 TYPE="TEXT" NAME="TxtCALLFLOW_ID" VALUE="<%=RSCALLFLOW_ID%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
	<td CLASS=LABEL>Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=80 size=50 TYPE="TEXT" NAME="TxtNAME" VALUE="<%=RSNAME%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</TR>
	<TR>
	<td CLASS=LABEL COLSPAN=2>Description:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=2000 size=80 TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</TR>
</TABLE>

<% Else %>

<DIV style="margin-top:170px;margin-left:170px" CLASS="LABEL">
<% If Request.QueryString("STATUS") <> "" Then %>
<%= Request.QueryString("STATUS")%><BR>
<%End If %>
No Address selected.
</DIV>
<% End If %>

</form>
</body>
</html>


