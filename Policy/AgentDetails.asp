<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\CheckSharedAgent.inc"-->
<%	Response.Expires = 0
	Response.Buffer = true
 %>
<!--#include file="..\lib\ZIP.inc"-->	
<%	

	OfficeTextLen = 50

	RSAHSID = Request.QueryString("AHSID")
	
	Dim SharedCount, SharedCountText, AID
	SharedCount = 0
	SharedCountText = "Ready"

	AID = Request.QueryString("AID")
	
	If AID <> "" Then
		If AID = "NEW" Then 
			SharedCount = 0
		Else
			SharedCount = CheckSharedAgent(CLng(AID),true,true,1,false,false,0)
		End If
	End If	

	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Agent Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
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
<script>
function CBranchSearchObj()
{
	this.BID = "";
	this.BIDOfficeName = "";
	this.Selected = false;	
}
var BranchSearchObj = new CBranchSearchObj();
var g_StatusInfoAvailable = false;

</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
sub window_onload
<%	
	IF Request.QueryString("MODE") = "RO" THEN %>
		SetScreenFieldsReadOnly true, "DISABLED"
<%	ELSE
		If AID <> "" Then
			if SharedCount <= 1 then %>
				document.all.ChkEdit.checked = true
				ChkEdit_OnClick
		<%	else %>
				document.all.ChkEdit.checked = false
				ChkEdit_OnClick
				SetStatusInfoAvailableFlag(true)	
			<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
				If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
					document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 Color='Maroon'>+</Font>"
			<%	Else %>
					document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>"
			<%  End If 
			end if
		End If	'AID <> ""
	END IF %>
End Sub

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"

	if document.all.ChkEdit.checked = true then
		SetScreenFieldsReadOnly false, "LABEL"
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScreenFieldsReadOnly true, "DISABLED"
		document.body.setAttribute "ScreenMode", "RO"				
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"	
end sub

Function AttachOffice
	BID = document.all.spanOID.innerText
	MODE = document.body.getAttribute("ScreenMode")
	If MODE = "RO" Then 
		Exit Function
	End If

	BranchSearchObj.BID = BID
	BranchSearchObj.Selected = false

	If BID = "" Then BID = "NEW"
	
	If BID = "NEW" And MODE = "RO" Then
		MsgBox "No Branch currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "../Branch/BranchMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_AGENT&BranchTypeFilter=CLAIMHANDLING&OID=" & OID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,BranchSearchObj ,"center"
	If BranchSearchObj.Selected = true Then
		If BranchSearchObj.BID <> document.all.spanOID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			document.all.spanOID.innerText = BranchSearchObj.BID
			document.all.spanBranchName.innerText = BranchSearchObj.BIDOfficeName
		end if
	end if
End Function


Function DetachOffice
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		spanOID.innerText = ""
		spanBranchName.innerText = ""
	end if
End Function

Sub UpdateOfficeFields
	If Len(OfficeSearchObj.OIDType) < <%=OfficeTextLen%> Then
		spanTYPE.innertext = OfficeSearchObj.OIDType
	Else
		spanTYPE.innertext = Mid ( OfficeSearchObj.OIDType, 1, <%=OfficeTextLen%>) & " ..."
	End If
	spanTYPE.title = OfficeSearchObj.OIDType

	If Len(OfficeSearchObj.OIDNumber) < <%=OfficeTextLen%> Then
		spanNUMBER.innertext = OfficeSearchObj.OIDNumber
	Else
		spanNUMBER.innertext = Mid ( OfficeSearchObj.OIDNumber, 1, <%=OfficeTextLen%>) & " ..."
	End If
	spanNUMBER.title = OfficeSearchObj.OIDNumber

	If Len(OfficeSearchObj.OIDState) < <%=OfficeTextLen%> Then
		spanSTATE.innertext = OfficeSearchObj.OIDState
	Else
		spanSTATE.innertext = Mid ( OfficeSearchObj.OIDState, 1, <%=OfficeTextLen%>) & " ..."
	End If
	spanSTATE.title = OfficeSearchObj.OIDState

	If Len(OfficeSearchObj.OIDZip) < <%=OfficeTextLen%> Then
		spanZIP.innertext = OfficeSearchObj.OIDZip
	Else
		spanZIP.innertext = Mid ( OfficeSearchObj.OIDZip, 1, <%=OfficeTextLen%>) & " ..."
	End If
	spanZIP.title = OfficeSearchObj.OIDZip

End Sub


Sub PostTo(strURL)
	FrmDetails.action = "AgentSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateAID(inAID)
	document.all.AID.value = inAID
	document.all.spanAID.innerText = inAID
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

Function GetAID
	if document.all.AID.value <> "NEW" then
		GetAID = document.all.AID.value
	else
		GetAID = ""
	end if 
End Function

Function GetAIDName
	GetAIDName = document.all.TxtName.value
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
	ValidateScreenData = true
End Function

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
	
	If document.all.AID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
	
	document.body.setAttribute "ScreenDirty","YES"
	document.all.AID.value = "NEW"
	ExeCopy = ExeSave
End Function


Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.AID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.AID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "AGENT_ID" & Chr(129) & document.all.AID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "AGENT_NUMBER" & Chr(129) & document.all.AGENTNo.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "AGENT_BRANCHNUM" & Chr(129) & document.all.TxtNumber.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATUS" & Chr(129) & document.all.TxtStatus.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TYPECD" & Chr(129) & document.all.TxtType.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME" & Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CITY" & Chr(129) & document.all.City.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ZIPCODE" & Chr(129) & document.all.Zip.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "STATE" & Chr(129) & document.all.State.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE" & Chr(129) & document.all.TxtPhone.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX" & Chr(129) & document.all.TxtFax.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAXCD" & Chr(129) & document.all.TxtFaxCd.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LAT"& Chr(129) & document.all.TxtLAT.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LON"& Chr(129) & document.all.TxtLON.value & Chr(129) & "1" & Chr(128)
		
		sResult = sResult & "BRANCH_ID"& Chr(129) & document.all.spanOID.innerText & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "AgentSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
			
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

Sub RefCountRpt_onclick()
	If document.all.SpanSharedCount.innerText > 0 Then
		If document.all.AID.value <> "" And document.all.AID.value <> "NEW" Then
			paramID = document.all.AID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedAgent=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
	Else
		MsgBox "Reference count is zero.",0,"FNSNetDesigner"	
	End If	
End Sub

Sub RefCountRpt_onmouseover()
	If document.all.SpanSharedCount.innerText > 0 Then
		document.all.RefCountRpt.style.cursor = "HAND"
	Else
		document.all.RefCountRpt.style.cursor = "DEFAULT"
	End If
End Sub

<!--#include file="..\lib\Help.asp"-->

</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<form Name="FrmDetails" METHOD="POST" ACTION="AgentSave.asp" TARGET="hiddenPage">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Agent Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" name="SearchAID" value="<%=Request.QueryString("SearchAID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchBranchNum" value="<%=Request.QueryString("SearchBranchNum")%>">
<input type="hidden" name="SearchState" value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchZip" value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchOfficeNumber" value="<%=Request.QueryString("SearchOfficeNumber")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="AID" value="<%=Request.QueryString("AID")%>">
<%	
AID	= CStr(Request.QueryString("AID"))
If AID <> "" Then
	If AID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT A.*, O.OFFICE_TYPE, O.OFFICE_NAME, O.OFFICE_NUMBER, O.STATE OSTATE, " &_ 
				"O.ZIP OZIP  FROM AGENT A, BRANCH O WHERE " &_
				"A.BRANCH_ID = O.BRANCH_ID(+) AND " &_
				"A.AGENT_ID = " & AID
				
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then

			RSAGENT_BRANCHNUM = RS("AGENT_BRANCHNUM")
			RSSTATUS = RS("STATUS")
			RSTYPECD = RS("TYPECD")			
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSAGENTNO = RS("AGENT_NUMBER")
			RSCITY = RS("CITY")			
			RSSTATE = RS("STATE")			
			RSZIPCODE = RS("ZIPCODE")			
			RSPHONE = RS("PHONE")			
			RSFAX = RS("FAX")			
			RSFAXCD = RS("FAXCD")			
			RSLAT = RS("LAT")			
			RSLON = RS("LON")			
			RSTYPE = RS("OFFICE_TYPE")
			RSBRANCHNAME = RS("OFFICE_NAME")
			RSNUMBER = RS("OFFICE_NUMBER")
			RSOSTATE = RS("OSTATE")
			RSOZIP = RS("OZIP")
			RSOID = RS("BRANCH_ID")
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
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
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
<input ScrnBtn="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>


<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td colspan="2">Agent ID:&nbsp;<span id="spanAID"><%=Request.QueryString("AID")%></span></td></tr>
<tr><td>Branch Number:<br><input type="text" ScrnInput="TRUE" size="7" class="LABEL" name="TxtNumber" maxlength="6" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSAGENT_BRANCHNUM%>"></td>
	<td>Status:<br><input type="text" ScrnInput="TRUE" size="2" class="LABEL" name="TxtStatus" maxlength="1" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSSTATUS%>"></td>
    <td>Type:<br><input type="text" ScrnInput="TRUE" size="2" class="LABEL" name="TxtType" maxlength="1" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSTYPECD%>"></td>
	<td COLSPAN="2">Name:<br><input type="text" ScrnInput="TRUE" size="32" class="LABEL" name="TxtName" maxlength="41" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSNAME%>"></td>    
	<td>Agent No.:<br><input type="text" ScrnInput="TRUE" size="12" class="LABEL" name="AGENTNo" maxlength="32" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSAGENTNO%>"></td>
</tr>
<tr>
    <td>Zip:<br><input type="text" ScrnInput="TRUE" size="10" class="LABEL" name="Zip" maxlength="9" value="<%=RSZIPCODE%>"></td>
	<td COLSPAN="3">City:<br><input type="text" size="32" class="READONLY" READONLY TABINDEX=-1 name="City" maxlength="41" value="<%=RSCITY%>"></td>
    <td>State:<br><input type="text" size="3" class="READONLY" READONLY TABINDEX=-1 name="STATE" maxlength="3" value="<%=RSSTATE%>"></td>
</tr>
<tr><td COLSPAN="3">Phone:<br><input type="text" ScrnInput="TRUE" size="32" class="LABEL" name="TxtPhone" maxlength="14" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSPHONE%>"></td>
    <td COLSPAN="2">Fax:<br><input type="text" ScrnInput="TRUE" size="32" class="LABEL" name="TxtFax" maxlength="10" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSFAX%>"></td>
    <td COLSPAN="2">Fax Code:<br><input type="text" ScrnInput="TRUE" size="2" class="LABEL" name="TxtFaxCd" maxlength="1" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSFAXCD%>"></td>
</tr>
<tr><td>LAT:<br><input type="text" ScrnInput="TRUE" size="14" class="LABEL" name="TxtLAT" maxlength="10" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSLAT%>"></td>
    <td COLSPAN="2">LON:<br><input type="text" ScrnInput="TRUE" size="14" class="LABEL" name="TxtLON" maxlength="10" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange" value="<%=RSLON%>"></td>
</tr>

</table>

<br>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Branch&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table class="Label">
<td>
<img NAME="BtnAttachOffice" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Office" ONCLICK="VBScript::AttachOffice">
<img NAME="BtnDetachOffice" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Office" OnClick="VBScript::DetachOffice">
</td>
<td>Branch ID:&nbsp;<span ID="spanOID" CLASS="LABEL"><%=RSOID%></span>
<span ID="spanBranchName" CLASS="LABEL"><%=RSBRANCHNAME%></span>
</td>
</table>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No agent selected.
</div>


<% End If %>

</form>
</body>
</html>


