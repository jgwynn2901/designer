<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\CheckSharedAttribute.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	Response.Expires=0 
		
	Dim SharedCount, SharedCountText, AID
	SharedCount = 0
	SharedCountText = "Ready"
	
	AID	= CStr(Request.QueryString("AID"))
	
	If AID <> "" Then
		If AID = "NEW" Then 
			SharedCount = 0
		Else
			SharedCount = CheckSharedAttribute(CLng(AID),True,True,1,False,False,0)
		End If
	End If	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Attribute Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT>
function CLookupTypeSearchObj()
{
	this.LUTID = "";
	this.LUTIDName = "";
	this.Selected = false;
}

function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}

var RuleSearchObj = new CRuleSearchObj();
var LookupTypeSearchObj = new CLookupTypeSearchObj();
var g_StatusInfoAvailable = false;


</SCRIPT>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<%	IF CStr(Request.QueryString("MODE")) = "RO" THEN %>
	SetScreenFieldsReadOnly true,"DISABLED"
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

Sub PostTo(strURL)
	FrmDetails.action = "AttributeSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub BtnAttachLookupType_OnClick
	LUTID = document.all.LU_TYPE_ID.value
	MODE = document.body.getAttribute("ScreenMode")

	LookupTypeSearchObj.LUTID = LUTID
	LookupTypeSearchObj.LUTIDName = document.all.LOOKUPNAME_TEXT.innerText
	LookupTypeSearchObj.Selected = false

	If LUTID = "" Then LUTID = "NEW"
	
	If LUTID = "NEW" And MODE = "RO" Then
		MsgBox "No Lookups currently attached.",0,"FNSNetDesigner"
		Exit Sub
	End If
	
	
	strURL = "..\LookupType\LookupTypeMaintenance.asp?SECURITYPRIV=FNSD_ATTRIBUTE&CONTAINERTYPE=MODAL&LUTID=" & LUTID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,LookupTypeSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If LookupTypeSearchObj.Selected = true Then
		If LookupTypeSearchObj.LUTID <> document.all.LU_TYPE_ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			document.all.LU_TYPE_ID.value = LookupTypeSearchObj.LUTID
		end if
		UpdateLookupText (document.all.LOOKUPNAME_TEXT)
	ElseIf document.all.LU_TYPE_ID.value = LookupTypeSearchObj.LUTID And LookupTypeSearchObj.LUTID <> "" Then
		UpdateLookupText (document.all.LOOKUPNAME_TEXT)
	End If
End Sub

Sub UpdateLookupText (SPANID)
	If Len(LookupTypeSearchObj.LUTIDName) < 22 Then
		SPANID.innertext = LookupTypeSearchObj.LUTIDName
	Else
		SPANID.innertext = Mid ( LookupTypeSearchObj.LUTIDName, 1, 22) & " ..."
	End If
	SPANID.title = LookupTypeSearchObj.LUTIDName
End Sub

Sub UpdateRuleText (SPANID)
	If Len(RuleSearchObj.RIDText) < 40 Then
		SPANID.innertext = RuleSearchObj.RIDText
	Else
		SPANID.innertext = Mid ( RuleSearchObj.RIDText, 1, 40) & " ..."
	End If
	SPANID.title = RuleSearchObj.RIDText
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
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ATTRIBUTE&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = RuleSearchObj.RID
		end if
		UpdateRuleText(SPANID)
	ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateRuleText(SPANID)
	End If

End Function


Function DetachRule(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.value = ""
		SPANID.innerText = ""
	end if
End Function

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

Function GetAIDCaption
	GetAIDCaption = document.all.TxtCaption.value
End Function

Function GetAIDInputType
	GetAIDInputType = document.all.TxtInputType.value
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
	
	if document.all.TxtLength.value <> "" then
		if IsNumeric(document.all.TxtLength.value) = false then
			MsgBox "Please enter a number in the Text Length field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	end if
	ValidateScreenData = true
End Function

sub UpdateScreenOnDelete()
	document.all.AID.value = ""
	FrmDetails.action = "AttributeDetails.asp?STATUS=Delete successful."
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
	
	if document.all.AID.value = "" then
		ExeDelete = false
		exit function
	end if

	if document.all.SpanSharedCount.innerText <> "0" Then
		MsgBox "Shared Count is " & document.all.SpanSharedCount.innerText & vbCRLF & _
		"You cannot delete an attribute with a shared count greater than zero.",vbExclamation,"FNSNetDesigner"
		ExeDelete = false
		exit Function
	end if

	lret = Confirm("Are you sure you want to delete Attribute ID: " & document.all.AID.value & " ?")

	if lRet = true Then
		document.all.TxtAction.value = "DELETE"
		sResult = document.all.AID.value
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
	
	if document.all.AID.value = "" then
		ExeCopy = false
		exit function
	end if


	document.all.AID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	document.all.SpanSharedCount.innerText = 0
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	if document.all.SpanSharedCount.innerText <> "0" Then
		lret = Confirm("Changes to the Attribute level will cause serious potential impact to other client's call flows in production.  Are you sure you want to change the Attribute level?")
		if not lRet Then
			ExeSave = bRet
			exit Function
		end	if 
	end if
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.AID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.AID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "ATTRIBUTE_ID"& Chr(129) & document.all.AID.value & Chr(129) & "1" & Chr(128)

		document.all.TxtName.value = Replace(document.all.TxtName.value," ","")		
		document.all.TxtName.value = UCase(document.all.TxtName.value)		
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		
		sResult = sResult & "CAPTION"& Chr(129) & document.all.TxtCaption.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ENTRYMASK"& Chr(129) & document.all.TxtEntryMask.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DEFAULTVALUE"& Chr(129) & document.all.TxtDefaultValue.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "UNKNOWNVALUE"& Chr(129) & document.all.TxtUnknownValue.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TEXTLENGTH"& Chr(129) & document.all.TxtLength.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "INPUTTYPE"& Chr(129) & document.all.TxtInputType.value & Chr(129) & "1" & Chr(128)

		if document.all.ChkValidValue.checked = True then
			sResult = sResult & "VALIDVALUEFIELD_FLG"& Chr(129) & "Y"  & Chr(129) & "1" & Chr(128)
		else 
			sResult = sResult & "VALIDVALUEFIELD_FLG"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if
		if document.all.ChkSpellCheck.checked = True then
			sResult = sResult & "SPELLCHECK_FLG"& Chr(129) & "Y"  & Chr(129) & "1" & Chr(128)
		else 
			sResult = sResult & "SPELLCHECK_FLG"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if
		
		sResult = sResult & "LU_TYPE_ID"& Chr(129) & document.all.LU_TYPE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "HELPSTRING"& Chr(129) & document.all.TxtHelpString.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "VISIBLERULE_ID"& Chr(129) & document.all.VISIBLE_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ENABLEDRULE_ID"& Chr(129) & document.all.ENABLED_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "VALIDRULE_ID"& Chr(129) & document.all.VALID_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PERSISTRULE_ID"& Chr(129) & document.all.PERSIST_ID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACTION_ID"& Chr(129) & document.all.ACTION_ID.value & Chr(129) & "1" & Chr(128)

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

Sub RefCountRpt_onclick()
	If document.all.SpanSharedCount.innerText > 0 Then
		If document.all.AID.value <> "" And document.all.AID.value <> "NEW" Then
			paramID = document.all.AID.value
		Else	
			paramID = 0
		End If
		lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedAttribute=True&ID=" & paramID, Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
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
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Attribute Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="AttributeSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchAID" value="<%=Request.QueryString("SearchAID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchCaption" value="<%=Request.QueryString("SearchCaption")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchHelpString" value="<%=Request.QueryString("SearchHelpString")%>">
<input type="hidden" name="SearchInputType" value="<%=Request.QueryString("SearchInputType")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="AID" value="<%=Request.QueryString("AID")%>" >

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

If AID <> "" Then
	If AID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ATTRIBUTE_LU_VIEW WHERE ATTRIBUTE_ID = " & AID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSCAPTION = ReplaceQuotesInText(RS("CAPTION"))
			RSTEXTLENGTH = RS("TEXTLENGTH")
			RSINPUTTYPE = RS("INPUTTYPE")
			RSENTRYMASK = ReplaceQuotesInText(RS("ENTRYMASK"))
			RSUNKNOWNVALUE = ReplaceQuotesInText(RS("UNKNOWNVALUE"))
			RSDEFAULTVALUE = ReplaceQuotesInText(RS("DEFAULTVALUE"))
			RSLU_NAME = RS("LU_NAME")
			RSLU_TYPE_ID = RS("LU_TYPE_ID")
			RSVALIDVALUEFIELD_FLG	= RS("VALIDVALUEFIELD_FLG")
			RSSPELLCHECK_FLG = RS("SPELLCHECK_FLG")
			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
			RSHELPSTRING = ReplaceQuotesInText(RS("HELPSTRING"))
			RSVISIBLERULE_TEXT = RS("VISIBLERULE_TEXT")
			RSVISIBLERULE_ID = RS("VISIBLERULE_ID")
			RSENABLEDRULE_TEXT = RS("ENABLEDRULE_TEXT")
			RSENABLEDRULE_ID = RS("ENABLEDRULE_ID")
			RSVALIDRULE_TEXT = RS("VALIDRULE_TEXT")
			RSVALIDRULE_ID = RS("VALIDRULE_ID")
			RSPERSISTRULE_TEXT = RS("PERSISTRULE_TEXT")
			RSPERSISTRULE_ID = RS("PERSISTRULE_ID")
			RSACTION_TEXT = RS("ACTION_TEXT")
			RSACTION_ID = RS("ACTION_ID")
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
%>
<table ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td WIDTH="14">
<img ID = "RefCountRpt" SRC="..\images\RefCount.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count">
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">
:<span id="SpanSharedCount"><%=SharedCount%></span>
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</td>
<td>
<input  ScrnBtn="TRUE"  TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>

<table CLASS="LABEL" CELLPADDING=0 CELLSPACING=0 id="TblControls">
<tr>
<td>
	<table class="LABEL">
	<tr>
	<tr>
	<tr>
	<tr>
<td>Attribute ID:&nbsp<span id="spanAID"><%=Request.QueryString("AID")%></span></td>
	<td>
	<td><input ScrnBtn="TRUE" TYPE="CHECKBOX" NAME="ChkValidValue" ONCLICK="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" <% If CStr(RSVALIDVALUEFIELD_FLG) = "Y" Then Response.Write("CHECKED")%>>Valid Value?</td>
	<td align=right><input ScrnBtn="TRUE" TYPE="CHECKBOX" NAME="ChkSpellCheck" ONCLICK="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" <% If CStr(RSSPELLCHECK_FLG) = "Y" Then Response.Write("CHECKED")%>>Spell Check?</td>
	</tr>
	<tr>
	<td COLSPAN=4>Name:<br><input style="text-transform:uppercase" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=255 SIZE=85 TYPE="TEXT" NAME="TxtName" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td COLSPAN=4>Caption:<br><input ScrnInput="TRUE" size=85 CLASS="LABEL" MAXLENGTH=80 TYPE="TEXT" NAME="TxtCaption" VALUE="<%=RSCAPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td>Unknown Value:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=255 size="20" TYPE="TEXT" NAME="TxtUnknownValue" VALUE="<%=RSUNKNOWNVALUE%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Input Type:<br><SELECT ScrnBtn = "TRUE" NAME="TxtInputType" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetValidValuesHTML("ATTRIBUTE_INPUTTYPE",RSINPUTTYPE,true)%></SELECT></td>
	<td>Text Length:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size="12" TYPE="TEXT" NAME="TxtLength" VALUE="<%=RSTEXTLENGTH%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Entry Mask:<br><input ScrnInput="TRUE" size="25" CLASS="LABEL" MAXLENGTH=80 TYPE="TEXT" NAME="TxtEntryMask" VALUE="<%=RSENTRYMASK%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td>Default Value:<br><input ScrnInput="TRUE" CLASS="LABEL" size="20" MAXLENGTH=2000 TYPE="TEXT" NAME="TxtDefaultValue" VALUE="<%=ReplaceRuleText(RSDEFAULTVALUE)%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td nowrap ONDRAGSTART="return false;" colspan=3>
	<table class="LABEL">
		<td nowrap>
		<IMG NAME=BtnAttachLookupType TITLE="Attach Lookup Type" SRC="..\IMAGES\Attach.gif">
		<IMG NAME=BtnDetachLookupType TITLE="Detach Lookup Type" OnClick="VBScript::DetachRule LU_TYPE_ID, LOOKUPNAME_TEXT" SRC="..\IMAGES\Detach.gif"></td>
		<td nowrap>Lookup Name:&nbsp<SPAN ID=LOOKUPNAME_TEXT CLASS=LABEL TITLE="<%=ReplaceRuleText(RSLU_NAME)%>" ><%=TruncateLookupText(RSLU_NAME)%></SPAN>
		<input type="hidden" name="LU_TYPE_ID" value="<%=RSLU_TYPE_ID%>"></td>
	<td>	
	<td>
	</table>	
	</td></tr>
	</table>
</td>
</tr> 
</table>

 
<table CLASS="LABEL" WIDTH=100%>
<tr><td>Description:<br><input ScrnInput="TRUE" MAXLENGTH=255 CLASS="LABEL" size="85" TYPE="TEXT" NAME="TxtDescription" VALUE="<%=RSDESCRIPTION%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
<tr><td>Help String:<br><input ScrnInput="TRUE" MAXLENGTH=2000 CLASS="LABEL" size="85" TYPE="TEXT" NAME="TxtHelpString" VALUE="<%=RSHELPSTRING%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
</table>
<table CLASS="LABEL" ONDRAGSTART="return false;">
<tr>
<td>
<IMG NAME=BtnAttachVisible STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule VISIBLE_ID, VISIBLE_TEXT,'Visible'">
<IMG NAME=BtnDetachVisible STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule VISIBLE_ID, VISIBLE_TEXT">
</TD>
<td nowrap>Visible Rule:</td>
<td><SPAN ID=VISIBLE_TEXT CLASS=LABEL TITLE="<%=ReplaceRuleText(RSVISIBLERULE_TEXT)%>"><%=TruncateRuleText(RSVISIBLERULE_TEXT)%></SPAN><INPUT TYPE=HIDDEN NAME=VISIBLE_ID VALUE="<%= RSVISIBLERULE_ID %>"></TD>
</tr>
<tr>
<td>
<IMG NAME=BtnAttachEnabled STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule ENABLED_ID, ENABLED_TEXT,'Enabled'">
<IMG NAME=BtnDetachEnabled STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule ENABLED_ID, ENABLED_TEXT">
</td>
<td nowrap>Enabled Rule:</td>
<td><SPAN ID=ENABLED_TEXT CLASS=LABEL TITLE="<%=ReplaceRuleText(RSENABLEDRULE_TEXT)%>" ><%=TruncateRuleText(RSENABLEDRULE_TEXT)%></SPAN><INPUT TYPE=HIDDEN NAME=ENABLED_ID VALUE="<%= RSENABLEDRULE_ID %>"></TD>
</tr>
<tr>
<td>
<IMG NAME=BtnAttachValid STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule VALID_ID, VALID_TEXT,'Valid'">
<IMG NAME=BtnDetachValid STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule VALID_ID, VALID_TEXT">
</td>
<td nowrap>Valid Rule:</td>
<td><SPAN ID=VALID_TEXT CLASS=LABEL TITLE="<%=ReplaceRuleText(RSVALIDRULE_TEXT)%>"><%=TruncateRuleText(RSVALIDRULE_TEXT)%></SPAN><INPUT TYPE=HIDDEN NAME=VALID_ID VALUE="<%= RSVALIDRULE_ID %>"></TD>
</tr>
<td>
<IMG NAME=BtnAttachPersist STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule PERSIST_ID, PERSIST_TEXT,'Persist'">
<IMG NAME=BtnDetachPersist STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule PERSIST_ID, PERSIST_TEXT">
</TD>
<td nowrap>Persist Rule:</td>
<td><SPAN ID=PERSIST_TEXT CLASS=LABEL TITLE="<%=ReplaceRuleText(RSPERSISTRULE_TEXT)%>"><%=TruncateRuleText(RSPERSISTRULE_TEXT)%></SPAN><INPUT TYPE=HIDDEN NAME=PERSIST_ID VALUE="<%= RSPERSISTRULE_ID %>"></TD>
</tr>
<td>
<IMG NAME=BtnAttachAction STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule ACTION_ID, ACTION_TEXT,'Action'">
<IMG NAME=BtnDetachAction STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule ACTION_ID, ACTION_TEXT">
<td nowrap>Action:</td>
<td><SPAN ID=ACTION_TEXT CLASS=LABEL TITLE="<%=ReplaceRuleText(RSACTION_TEXT)%>" ><%=TruncateRuleText(RSACTION_TEXT)%></SPAN><INPUT TYPE=HIDDEN NAME=ACTION_ID VALUE="<%= RSACTION_ID %>"></TD>
</tr>
</table>

<% Else %>

<DIV style="margin-top:170px;margin-left:170px" CLASS="LABEL">
<%=Request.QueryString("STATUS") & "<br>"%>
No attribute selected.
</DIV>


<% End If %>

</form>
</body>
</html>


