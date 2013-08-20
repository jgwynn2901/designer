<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<% Response.Expires=0 
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
If Request.QueryString("ATTRIBUTEOVERRIDE_ID") <> "NEW" Then
	SQL = "SELECT * FROM ATTRIBUTE_OVERRIDE WHERE ATTRIBUTEOVERRIDE_ID=" & Request.QueryString("ATTRIBUTEOVERRIDE_ID")
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
	PROPERTY_NAME = RS("PROPERTY_NAME")
	SEQUENCE = RS("SEQUENCE")
	LU_TYPE_ID = RS("LU_TYPE_ID")
	If IsNull(RS("CAPTION")) Then
		CAPTION = RS("CAPTION")
	Else
		CAPTION = Replace(RS("CAPTION"), """", """""")
	End If
	If IsNull(RS("INPUTTYPE")) Then
		INPUTTYPE = RS("INPUTTYPE")
	Else
		INPUTTYPE = Replace(RS("INPUTTYPE"), """", """""")
	End If
	If IsNull(RS("ENTRYMASK")) Then
		ENTRYMASK = RS("ENTRYMASK")
	Else
		ENTRYMASK = Replace(RS("ENTRYMASK"), """", """""")
	End If
	If IsNull(RS("VALIDVALUEFIELD_FLG")) Then
		VALIDVALUEFIELD_FLG = RS("VALIDVALUEFIELD_FLG")
	Else
		VALIDVALUEFIELD_FLG = Replace(RS("VALIDVALUEFIELD_FLG"), """", """""")
	End If
	If IsNull(RS("DEFAULTVALUE")) Then
		DEFAULTVALUE = RS("DEFAULTVALUE")
	Else
		DEFAULTVALUE = Replace(RS("DEFAULTVALUE"), """", """""")
	End If
	If IsNull(RS("UNKNOWNVALUE")) Then
		UNKNOWNVALUE = RS("UNKNOWNVALUE")
	Else
		UNKNOWNVALUE = Replace(RS("UNKNOWNVALUE"), """", """""")
	End If
	TEXTLENGTH = RS("TEXTLENGTH")
	ENABLEDRULE_ID = RS("ENABLEDRULE_ID")
	VALIDRULE_ID = RS("VALIDRULE_ID")
	PERSISTRULE_ID = RS("PERSISTRULE_ID")
	ACTION_ID = RS("ACTION_ID")
	If IsNull(RS("SPELLCHECK_FLG")) Then
		SPELLCHECK_FLG = RS("SPELLCHECK_FLG")
	Else
		SPELLCHECK_FLG = Replace(RS("SPELLCHECK_FLG"), """", """""")
	End If
	If IsNull(RS("HELPSTRING")) Then
		HELPSTRING = RS("HELPSTRING")
	Else
		HELPSTRING = Replace(RS("HELPSTRING"), """", """""")
	End If
	If IsNull(DESCRIPTION = RS("DESCRIPTION")) Then
		DESCRIPTION = RS("DESCRIPTION")
	Else
		DESCRIPTION = Replace(RS("DESCRIPTION"), """", """""")
	End If
	OVERRIDE_RULE_ID = RS("OVERRIDE_RULE_ID")
	RS.Close 
	SQL = "SELECT * FROM RULES WHERE RULE_ID=" & OVERRIDE_RULE_ID
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
	RS_RuleText = replace(trim(RS.Fields("RULE_TEXT").Value), """", """""")
	RS.Close 
End If
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
<!--#include file="..\lib\Help.asp"-->

Sub PROPERTY_NAME_onchange
window.setTimeout "SetOverride()", 10
End Sub

Sub window_onload
	<% If Request.QueryString("ATTRIBUTEOVERRIDE_ID") = "NEW" Then %>
		SpanStatus.innerHTML = "New"
		PROPERTY_NAME.value = ""
		document.all.TEXTBOXVALUE.style.display = "none"
		document.all.FLAGVALUE.style.display = "none"
		document.all.OVERRIDELABEL.innerHTML = "New"
		document.all.OVERRIDE_RULE_ID.value = ""
	<% Else %>
		SpanStatus.innerHTML = "Ready"
		PROPERTY_NAME.value = "<%= PROPERTY_NAME %>"
		window.setTimeout  "SetOverride()", 500
		document.all.OVERRIDE_RULE_ID.value = "<%= OVERRIDE_RULE_ID%>"
		document.all.ruleText.innerHTML = "<%=RS_RuleText%>"
	<% End If %>
	<% If PROPERTY_NAME = "VALIDVALUEFIELD_FLG" AND VALIDVALUEFIELD_FLG= "Y" Then %>
		document.all.TEXTBOXVALUE.style.display = "none"
		document.all.FLAGVALUE.style.display = "block"
		document.all.OVERRIDEFLAG.checked = true
	<% End If %>
	<% If PROPERTY_NAME = "SPELLCHECK_FLG" AND SPELLCHECK_FLG= "Y" Then %>
		document.all.TEXTBOXVALUE.style.display = "none"
		document.all.FLAGVALUE.style.display = "block"
		document.all.OVERRIDEFLAG.checked = true
	<% End If %>
End Sub


Sub Switch()
	document.all.TEXTBOXVALUE.style.display = "block"
	document.all.FLAGVALUE.style.display = "none"
End Sub

Sub SetRead()
	document.all.ATTACHBUTTONS.style.display = "block"
	document.all.OVERRIDEVALUE.readonly = true
	document.all.OVERRIDEVALUE.style.backgroundcolor = "Silver"
End Sub

Sub SetWrite()
	document.all.ATTACHBUTTONS.style.display = "none"
	document.all.OVERRIDEVALUE.readonly = false
	document.all.OVERRIDEVALUE.style.backgroundcolor = "White"
End Sub

Function SetOverride()
	document.all.OVERRIDELABEL.innerHTML = PROPERTY_NAME(PROPERTY_NAME.selectedIndex).Text
	Select Case PROPERTY_NAME.value
		Case "CAPTION"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 80
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= CAPTION %>"
		Case "INPUTTYPE"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 80
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= INPUTTYPE %>"
		Case "ENTRYMASK"	
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 80
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= ENTRYMASK%>"
		Case "DEFAULTVALUE"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 255
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= DEFAULTVALUE%>"
		Case "UNKNOWNVALUE"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 255
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= UNKNOWNVALUE%>"
		Case "TEXTLENGTH"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 10
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= TEXTLENGTH %>"
		Case "DESCRIPTION"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.maxLength = 255
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.value = "<%= DESCRIPTION%>"
		Case "ACTION_ID"
			Switch()
			setRead()
			document.all.OVERRIDEVALUE.value = "<%= ACTION_ID%>"
			document.all.OVERRIDEVALUE.maxLength = 10
			document.all.OVERRIDEVALUE.size = 10
		Case "PERSISTRULE_ID" 
			Switch()
			setRead()
			document.all.OVERRIDEVALUE.value = "<%= PERSISTRULE_ID%>"
			document.all.OVERRIDEVALUE.maxLength = 10
			document.all.OVERRIDEVALUE.size = 10
		Case "ENABLEDRULE_ID"
			Switch()
			setRead()
			document.all.OVERRIDEVALUE.value = "<%=ENABLEDRULE_ID %>"
			document.all.OVERRIDEVALUE.maxLength = 10
			document.all.OVERRIDEVALUE.size = 10
		Case "VALIDRULE_ID"
			Switch()
			setRead()
			document.all.OVERRIDEVALUE.value = "<%= VALIDRULE_ID%>"
			document.all.OVERRIDEVALUE.maxLength = 10
			document.all.OVERRIDEVALUE.size = 10
		Case "LU_TYPE_ID"
			Switch()
			setRead()
			document.all.OVERRIDEVALUE.value = "<%= LU_TYPE_ID %>"
			document.all.OVERRIDEVALUE.maxLength = 10
			document.all.OVERRIDEVALUE.size = 10
		Case "HELPSTRING"
			Switch()
			SetWrite()
			document.all.OVERRIDEVALUE.size = 100
			document.all.OVERRIDEVALUE.maxLength = 2000
			document.all.OVERRIDEVALUE.value = "<%= HELPSTRING%>"
		Case "SPELLCHECK_FLG"
			document.all.ATTACHBUTTONS.style.display = "none"
			document.all.TEXTBOXVALUE.style.display = "none"
			document.all.FLAGVALUE.style.display = "block"
			<% If SPELLCHECK_FLG = "Y" Then %>
				document.all.OVERRIDEFLAG.checked = true
			<% End If %>
		Case "VALIDVALUEFIELD_FLG"
			document.all.ATTACHBUTTONS.style.display = "none"
			document.all.TEXTBOXVALUE.style.display = "none"
			document.all.FLAGVALUE.style.display = "block"
			<% If SPELLCHECK_FLG = "Y" Then %>
				document.all.OVERRIDEFLAG.checked = true
			<% End If %>
		Case Else
			PROPERTY_NAME.value = ""
			document.all.TEXTBOXVALUE.style.display = "none"
			document.all.FLAGVALUE.style.display = "none"
			document.all.OVERRIDELABEL.innerHTML = "New"
			document.all.OVERRIDE_RULE_ID.value = ""
	End Select
End Function

Sub BtnGrfxBack_OnClick()
	location.href = "Attribute_Override_Details.asp?<%= Request.QueryString %>"
End Sub

Sub BtnSave_onclick
errmsg = ""
If PROPERTY_NAME.value = "" Then
	errmsg = errmsg & "Please choose a Property Name." & vbcrlf
End If
If SEQUENCE.value = ""  Then
	errmsg = errmsg & "Sequence is a required field." & vbcrlf
End If
If Not isnumeric(SEQUENCE.value) Then
	errmsg = errmsg & "Sequence must be numeric." & vbcrlf
End If
If document.all.OVERRIDE_RULE_ID.value = "" Then
	errmsg = errmsg & "Over Ride Rule is a required field." & vbcrlf
End If

<% If Request.Querystring("ATTRIBUTEOVERRIDE_ID") = "NEW" Then %>
document.all.SAVEACTION.value = "INSERT"
<% Else %>
document.all.SAVEACTION.value = "UPDATE"
<% End If %>

OverrideValue = document.all.OVERRIDEVALUE.value
If document.all.PROPERTY_NAME.value = "SPELLCHECK_FLG" OR document.all.PROPERTY_NAME.value = "VALIDVALUEFIELD_FLG" Then
	If document.all.OVERRIDEFLAG.checked = True Then
		OverrideValue = "Y"
	Else
		OverrideValue = "N"
	End If
End If
	sResult = ""
	For i = 0 to PROPERTY_NAME.length-1 step 1
		If PROPERTY_NAME.selectedIndex = i Then
			sResult = sResult & document.all.PROPERTY_NAME.value & Chr(129) & OverrideValue & Chr(129) & "1" & Chr(128)
		Else
			sResult = sResult & document.all.PROPERTY_NAME(i).value & Chr(129) & "" & Chr(129) & "1" & Chr(128)
		End If
	Next
	
	sResult = sResult & "PROPERTY_NAME" & Chr(129) & document.all.PROPERTY_NAME.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ATTRIBUTEOVERRIDE_ID" & Chr(129) & document.all.ATTRIBUTEOVERRIDE_ID.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ATTR_INSTANCE_ID" & Chr(129) & "<%= Request.QueryString("ATTR_INSTANCE_ID") %>" & Chr(129) & "1" & Chr(128)
	sResult = sResult & "OVERRIDE_RULE_ID" & Chr(129) & document.all.OVERRIDE_RULE_ID.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "SEQUENCE" & Chr(129) & document.all.SEQUENCE.Value & Chr(129) & "1" & Chr(128)
	document.all.SAVEDATA.value = sResult
	If errmsg = "" Then
		FrmSave.submit()
	Else
		Msgbox errmsg,0, "FNSDesigner"
	End If
End Sub

Sub BtnCancel_onclick
	BtnGrfxBack_OnClick()
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Sub SetSpan(msg)
	SpanStatus.innerHTML = msg
End Sub

Sub UpdateStatus(msg)
	INST_ID.innerHTML = msg
End Sub

Function AttachRule ()
	RID = document.all.OVERRIDEVALUE.value
	RuleSearchObj.RID = RID
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	if document.all.PROPERTY_NAME.value = "LU_TYPE_ID" Then
		strURL = "..\LookupType\LookupTypeMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&LUTID=" & RID 
		showModalDialog  strURL  ,LookupTypeSearchObj ,"center"
	Else
		strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID
		'strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&RID=" & RID &  "&TITLE=" & strTITLE & "&MODE=" & MODE 
		showModalDialog  strURL  ,RuleSearchObj ,"center"
	End If
		
		If RuleSearchObj.RID <> "" then
			document.all.OVERRIDEVALUE.value = RuleSearchObj.RID
		end if
		If LookupTypeSearchObj.LUTID <> "" then
			document.all.OVERRIDEVALUE.value = LookupTypeSearchObj.LUTID
		end if

End Function

Function AttachOverrideRule ()
	RID = document.all.OVERRIDE_RULE_ID.value
	RuleSearchObj.RID = RID
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID &  "&TITLE=" & strTITLE & "&MODE=" & MODE 

	showModalDialog  strURL  ,RuleSearchObj ,"center"
		If RuleSearchObj.RID <> "" then
			document.all.OVERRIDE_RULE_ID.value = RuleSearchObj.RID
			document.all.ruleText.innerHTML = RuleSearchObj.RIDText
		end if
End Function

Sub BtnAttach_onclick
	AttachRule()
End Sub

Sub BtnOverrideAtt_OnClick
	AttachOverrideRule()
End Sub

Sub BtnDetach_onclick()
	document.all.OVERRIDEVALUE.value = ""
End Sub

Sub BtnNew_OnClick()
	self.location.href = "Attribute_Override.asp?ATTRIBUTEOVERRIDE_ID=NEW&ATTR_INSTANCE_ID=<%= Request.QueryString("ATTR_INSTANCE_ID") %>"
End Sub
-->
</SCRIPT>
<SCRIPT LANGUAGE=JavaScript>
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CLookupTypeSearchObj()
{
	this.LUTID = "";
	this.LUTIDName = "";
	this.Selected = false;
}
var RuleSearchObj = new CRuleSearchObj();
var LookupTypeSearchObj = new CLookupTypeSearchObj();
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=3 topmargin=0 class=LABEL>
<!--#include file="..\lib\NavBack.inc"-->
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Attribute Override Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>

<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="4">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</TD>
<TD CLASS=LABEL>
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</td>
</tr>
</table>
<LABEL CLASS=LABEL>Attribute Override ID: <SPAN ID=INST_ID CLASS=LABEL><%= Request.Querystring("ATTRIBUTEOVERRIDE_ID")%></SPAN></LABEL>
<TABLE width="712" cellspacing="1" border="0" bordercolor="#000000" ID="Table2">
<TR>
<TD CLASS=LABEL width="125">Property Name:<br><SELECT ScrnBtn = "TRUE" NAME="PROPERTY_NAME" CLASS="LABEL" ID="Select2">
<OPTION VALUE="CAPTION">Caption
<OPTION VALUE="INPUTTYPE">Input Type
<OPTION VALUE="ENTRYMASK">Entry Mask
<OPTION VALUE="LU_TYPE_ID">LU Type ID
<OPTION VALUE="VALIDVALUEFIELD_FLG">Valid Value Field Flag
<OPTION VALUE="DEFAULTVALUE">Default Value
<OPTION VALUE="UNKNOWNVALUE">Unknown Value
<OPTION VALUE="TEXTLENGTH">Text Length
<OPTION VALUE="ENABLEDRULE_ID">Enabled Rule ID
<OPTION VALUE="VALIDRULE_ID">Valid Rule ID
<OPTION VALUE="PERSISTRULE_ID">Persist Rule ID
<OPTION VALUE="ACTION_ID">Action ID
<OPTION VALUE="SPELLCHECK_FLG">Spell Check Flag
<OPTION VALUE="HELPSTRING">Helpstring
<OPTION VALUE="DESCRIPTION">Description
</SELECT></TD>
<TD CLASS=LABEL width="538">Sequence:<BR>
<INPUT TYPE=TEXT NAME=SEQUENCE SIZE=4 MAXLENGTH=4 CLASS=LABEL VALUE="<%=SEQUENCE%>" ID="Text4">
</TD>
</tr>
</table>
<table ID="Table3">
<tr>
<TD CLASS=LABEL width="90">Override Rule:<BR>
<IMG ID="Img2" NAME=BtnOverrideAtt STYLE="cursor:hand" SRC="..\images\Attach.gif" TITLE="Attach Rule" width="16" height="16">
<INPUT CLASS=LABEL TYPE=TEXT SIZE=10 MAXLENGTH=10 STYLE="BACKGROUND-COLOR:SILVER" READONLY NAME=OVERRIDE_RULE_ID ID="Text5">
</TD>
<td CLASS=LABEL>
<div id=ruleText></div>
</td>
</TR>
</TABLE>
<BR>
<FORM NAME=FrmSave METHOD=POST ACTION="Attribute_Override_Save.asp" TARGET="hiddenPage">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Attribute Override Data</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<INPUT TYPE=HIDDEN NAME=SAVEACTION>
<INPUT TYPE=HIDDEN NAME=SAVEDATA>
<INPUT TYPE=HIDDEN NAME=ATTRIBUTEOVERRIDE_ID VALUE="<%= Request.Querystring("ATTRIBUTEOVERRIDE_ID")%>">
<FIELDSET ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'40%';width:'100%'">
<TABLE>
<TR>
<TD CLASS=LABEL>
<SPAN ID=OVERRIDELABEL CLASS=LABEL STYLE="DISPLAY:BLOCK"></SPAN>:<BR>
<SPAN ID="ATTACHBUTTONS" STYLE="DISPLAY:NONE">
<IMG ID=BtnAttach NAME=BtnAttach STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule">
<IMG ID=BtnDetach NAME=BtnDetach STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule">
</SPAN>


<SPAN ID=TEXTBOXVALUE CLASS=LABEL STYLE="DISPLAY:NONE">
<INPUT TYPE=TEXT NAME=OVERRIDEVALUE SIZE=100 CLASS=LABEL>
</SPAN>

<SPAN ID=FLAGVALUE CLASS=LABEL STYLE="DISPLAY:NONE">
<INPUT TYPE=CHECKBOX NAME=OVERRIDEFLAG CLASS=LABEL>
</SPAN>
</TD>
</TR>
</TABLE>
</FIELDSET>
<TABLE>
<TR>
<TD><BUTTON NAME=BtnSave CLASS=STDBUTTON ACCESSKEY="S"><U>S</U>ave</BUTTON></TD>
<TD><BUTTON NAME=BtnNew CLASS=STDBUTTON ACCESSKEY="N"><U>N</U>ew</BUTTON></TD>
<TD><BUTTON NAME=BtnCancel CLASS=STDBUTTON ACCESSKEY="B"><U>B</U>ack</BUTTON></TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
