<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	Response.Expires = 0 
	Response.Buffer = true
	AccountTextLen = 30	
	RSAHSID = Request.QueryString("AHSID")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Output Overflow Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
function CAttributeSearchObj()
{
	this.AID = "";
	this.AIDName = "";
	this.Selected = false;	
}
var AttributeSearchObj = new CAttributeSearchObj();
var AHSSearchObj = new CAHSSearchObj();

var g_StatusInfoAvailable = false;

</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub PostTo(strURL)
	FrmDetails.action = "OutputOverflowSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateOOID(inOOID)
	document.all.OOID.value = inOOID
	document.all.spanOOID.innerText = inOOID
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

Function GetOOID
	if document.all.OOID.value <> "NEW" then
		GetOOID = document.all.OOID.value
	else
		GetOOID = ""
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
errStr = ""

	If  document.all.TxtLOBCD.value = "" then errStr = "LOB is a required field." & VBCRLF
	If  document.all.AHSID_ID.innerText = "" then errStr = errStr & "A.H. Step ID is a required field." & VBCRLF
	If  document.all.TxtName.value = "" then errStr =  errStr & "Attribute Name is a required field." & VBCRLF
	If  document.all.TxtSequence.value = "" then 
		errStr =  errStr & "Sequence is a required field." & VBCRLF
	elseif IsNumeric(document.all.TxtSequence.value) = false then
			errStr =  errStr & "Sequence must be numeric." & VBCRLF
	end if
	
	if document.all.TxtLength.value <> "" then
		if IsNumeric(document.all.TxtLength.value) = false then errStr =  errStr & "Please enter a number in the Caption Length field." & VBCRLF
	end if

	If errstr = "" Then
		ValidateScreenData = true
	Else
		MsgBox errstr, 0 , "FNSNetDesigner"
		ValidateScreenData = false
	End If	
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
	
	If document.all.OOID.value = "" Then
		ExeCopy = false
		Exit Function
	End If
	
	document.body.setAttribute "ScreenDirty","YES"
	document.all.OOID.value = "NEW"
	ExeCopy = ExeSave
End Function


Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.OOID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if

		If document.all.OOID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if

		sResult = sResult & "OUTPUT_OVERFLOW_ID"& Chr(129) & document.all.OOID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD"& Chr(129) & document.all.TxtLOBCD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ATTRIBUTE_NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "SEQUENCE"& Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CAPTION"& Chr(129) & document.all.TxtCaption.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "MAPPING"& Chr(129) & document.all.TxtMapping.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "CAPTION_LENGTH"& Chr(129) & document.all.TxtLength.value & Chr(129) & "1" & Chr(128)

		if document.all.ChkShowWhenEmptyFlag.checked = True then
			sResult = sResult & "SHOW_WHEN_EMPTY_FLAG"& Chr(129) & "Y"  & Chr(129) & "1" & Chr(128)
		else 
			sResult = sResult & "SHOW_WHEN_EMPTY_FLAG"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if

		document.all.TxtSaveData.Value = sResult
		FrmDetails.action = "OutputOverflowSave.asp"
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
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_OUTPUT_OVERFLOW&SELECTONLY=TRUE&AHSID=" &AHSID
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

Function Detach(ID, SPANID)
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



Function LookupAttributeName (TEXTID)
	MODE = document.body.getAttribute("ScreenMode")

	if MODE =  "RO" then
		Exit Function
	End if
	
	AttributeSearchObj.AID = ""
	AttributeSearchObj.AIDName = TEXTID.value
	AttributeSearchObj.Selected = false

	strURL = "..\Attribute\AttributeMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_OUTPUT_OVERFLOW&SEARCHONLY=TRUE"
	
	showModalDialog  strURL  ,AttributeSearchObj ,"center"

	'if Selected=true update everything
	If AttributeSearchObj.Selected = true Then
		If AttributeSearchObj.AIDName <> TEXTID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			TEXTID.value = AttributeSearchObj.AIDName
		end if
	End If

End Function
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Output Overflow Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="OutputOverflowSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchOOID" value="<%=Request.QueryString("SearchOOID")%>">
<input type="hidden" name="SearchLOBCD" value="<%=Request.QueryString("SearchLOBCD")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchAttributeName" value="<%=Request.QueryString("SearchAttributeName")%>">
<input type="hidden" name="SearchSequence" value="<%=Request.QueryString("SearchSequence")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="OOID" value="<%=Request.QueryString("OOID")%>">

<%	
Dim OOID
OOID	= CStr(Request.QueryString("OOID"))
If OOID <> "" Then
	If OOID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT OUTPUT_OVERFLOW.*, ACCOUNT_HIERARCHY_STEP.NAME FROM OUTPUT_OVERFLOW, ACCOUNT_HIERARCHY_STEP " &_
				" WHERE OUTPUT_OVERFLOW.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+) AND " &_
				" OUTPUT_OVERFLOW_ID = " & OOID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
			RSAHSID_TEXT = RS("NAME")
			RSLOBCD = RS("LOB_CD")
			RSATTRIBUTE_NAME = ReplaceQuotesInText(RS("ATTRIBUTE_NAME")			)
			RSSEQUENCE = RS("SEQUENCE")			
			RSSHOW_WHEN_EMPTY_FLAG = RS("SHOW_WHEN_EMPTY_FLAG")			
			RSCAPTION = ReplaceQuotesInText(RS("CAPTION")			)
			RSCAPTION_LENGTH = RS("CAPTION_LENGTH")			
			RSMAPPING = ReplaceQuotesInText(RS("MAPPING"))
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<table class="LABEL">
<tr>
	<td>
	<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	</td>
	<td width=305 nowrap>Account:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(RSAHSID_TEXT)%>" ><%=TruncateText(RSAHSID_TEXT,AccountTextLen)%></SPAN></td>
	<td>A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL><%=RSAHSID%></SPAN></td>
	</tr>
</table>

<table CLASS="LABEL" >
<tr></tr>
<tr></tr>
<tr></tr> 
<tr></tr>
<tr><td colspan=3>Output Overflow ID:&nbsp;<span id="spanOOID"><%=Request.QueryString("OOID")%></span></td></tr>
<tr>
	<td>LOB:<br><select ScrnBtn="TRUE" name="TxtLOBCD" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RSLOBCD,true)%></select></td>
	<td><input ScrnBtn="TRUE" TYPE="CHECKBOX" NAME="ChkShowWhenEmptyFlag" ONCLICK="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" <% If CStr(RSSHOW_WHEN_EMPTY_FLAG) = "Y" Then Response.Write("CHECKED")%>>Show when empty?</td>
</tr>
<tr> 
	<td colspan=3>Attribute Name:<br>
	<IMG NAME=BtnLookupAttributeName STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Lookup Attribute Name" ONCLICK="VBScript::LookupAttributeName TxtName">
	<input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="60" TYPE="TEXT" NAME="TxtName" VALUE="<%=RSATTRIBUTE_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Sequence:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtSequence" VALUE="<%=RSSEQUENCE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
	<td colspan=3>Caption:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="60" TYPE="TEXT" NAME="TxtCaption" VALUE="<%=RSCAPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td>Caption Length:<br><input ScrnInput="TRUE" MAXLENGTH="10" CLASS="LABEL" size="12" TYPE="TEXT" NAME="TxtLength" VALUE="<%=RSCAPTION_LENGTH%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>
	<td colspan=4>Mapping:<br>
	<TEXTAREA ScrnBtn=TRUE COLS=80 ROWS=6 CLASS="LABEL" NAME="TxtMapping" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"><%=ReplaceQuotesInText(RSMAPPING)%></TEXTAREA>
</tr>
</table> 
 
<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No output overflow selected.
</div>


<% End If %>

</form>
</body>
</html>


