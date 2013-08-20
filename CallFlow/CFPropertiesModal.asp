<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE>Properties</TITLE>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT LANGUAGE=VBSCRIPT>
<!--
function PageCheck
StrError = ""

If Len(PropNAME.value) < 1 Then
	StrError = StrError & "Name is a required field" & VbCrlf
End if

If textlength.value <> "" AND Not Isnumeric(textlength.value) Then
StrError = StrError & "Text Length must be numeric" & VbCrlf
End If

If Len(PropWIDTH.value) < 1 OR Not IsNumeric(PropWIDTH.value) Then
	StrError = StrError & "Width is a required field, and must be numeric" & VbCrlf
End If

If Len(PropHEIGHT.value) < 1 OR Not IsNumeric(PropHEIGHT.value) Then
	StrError = StrError & "Height is a required field, and must be numeric" & VbCrlf
End If

If Len(PropXPOS.value) < 1 OR Not Isnumeric(PropXPOS.value) Then
	StrError = StrError & "XPos is a required field, and must be numeric" & VbCrlf
End If

If Len(PropYPOS.value) < 1 OR Not Isnumeric(PropYPOS.value) Then
	StrError = StrError & "YPos is a required field, and must be numeric" & VbCrlf
End If

If Len(PropSEQUENCE.value) < 1 OR Not Isnumeric(PropSEQUENCE.value) Then
	StrError = StrError & "Sequence is a required field, and must be numeric" & VbCrlf
End If

If StrError <> "" Then
	MsgBox StrError, 0 , "FNSDesigner"
	PageCheck = false
else
	PageCheck = true
end if

end function
-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
var inObj;
function window.onload()
{
		inObj  = window.dialogArguments;
		PropXPOS.value = inObj.xpos;
		PropYPOS.value = inObj.ypos;
		PropWIDTH.value = inObj.width;
		PropHEIGHT.value = inObj.height;
		PropNAME.value = inObj.name;
		PropSAMPLEVALUE.value = inObj.SampleValue;
		PropSEQUENCE.value = inObj.sequence;
		LUCOLUMN_NAME.value = inObj.lucolumn_name;
		LUSTORAGE_NAME.value = inObj.lustorage_name;
		PropTYPE.value = inObj.type;
//New
		

	
	PropCaption.value = inObj.caption
	if (inObj.caption == "-999999999")
	{PropCaption.style.color = "Maroon"}
	
	inputtype.value = inObj.inputtype
	if (inObj.inputtype == "-999999999") 
	{
		inputtype(inputtype.selectedIndex).style.color = "Maroon"
	}
	
	entrymask.value = inObj.entrymask
	if (inObj.entrymask == "-999999999")
	{entrymask.style.color = "Maroon"}
	
	if (inObj.validvaluefield_flg=="U") 
	{
	UNKNOWNVALIDVALUE.innerHTML = "*"
	UNKNOWNVALIDVALUE.style.color = "maroon"
	}
	else
	{
		if (inObj.validvaluefield_flg=="Y") 
		{validvaluefield_flg.checked = true;}
	}
	
	defaultvalue.value = inObj.defaultvalue
	if ( inObj.defaultvalue == "-999999999")
	{defaultvalue.style.color = "Maroon"}
	
	unknownvalue.value = inObj.unknownvalue
	if (inObj.unknownvalue == "-999999999")
	{unknownvalue.style.color = "Maroon"}
	
	textlength.value = inObj.textlength
	if ( inObj.textlength == "-999999999") 
	{textlength.style.color = "Maroon"}
	

		if (inObj.attributeframe_id == "" )
		{
			document.all.ATTRIBUTEFRAME_ID.value = "null"
			document.all.ATTRIBUTEFRAME_TEXT.innerHTML = ""
		}
		else
		{
		document.all.ATTRIBUTEFRAME_ID.value = inObj.attributeframe_id
		document.all.ATTRIBUTEFRAME_TEXT.innerHTML = (inObj.attributeframe_id == "null") ? "" : inObj.attributeframe_id
		}

	
	if (inObj.visiblerule_id == "-999999999")
	{
		document.all.VISIBLE_ID.value = "-999999999"
		document.all.VISIBLE_TEXT.innerHTML = "*Using attribute defined value*"
		document.all.VISIBLE_TEXT.title = "*Using attribute defined value*"
		document.all.VISIBLE_TEXT.style.color = "Maroon"
	}
	else
	{
		if (inObj.visiblerule_id == "") 
		{
		document.all.VISIBLE_ID.value = "null"
		document.all.VISIBLE_TEXT.innerHTML = ""
		}
		else
		{
		document.all.VISIBLE_ID.value = inObj.visiblerule_id
		document.all.VISIBLE_TEXT.innerHTML = (inObj.visiblerule_id == "null") ? "" : inObj.visiblerule_id
		}
	}

	if (inObj.enabledrule_id == "-999999999")
	{
		document.all.ENABLED_ID.value = "-999999999"
		document.all.ENABLED_TEXT.innerHTML = "*Using attribute defined value*"
		document.all.ENABLED_TEXT.title = "*Using attribute defined value*"
		document.all.ENABLED_TEXT.style.color = "Maroon"
	}
	else
	{
		if (inObj.enabledrule_id == "")
		{
		document.all.ENABLED_ID.value = "null"
		document.all.ENABLED_TEXT.innerHTML = ""
		}
		else
		{
		document.all.ENABLED_ID.value = inObj.enabledrule_id
		document.all.ENABLED_TEXT.innerHTML = (inObj.enabledrule_id == "null") ? "" : inObj.enabledrule_id
		}
	}

	if (inObj.validrule_id == "-999999999")
	{
		document.all.VALID_ID.value = "-999999999"
		document.all.VALID_TEXT.innerHTML = "*Using attribute defined value*"
		document.all.VALID_TEXT.title = "*Using attribute defined value*"
		document.all.VALID_TEXT.style.color = "Maroon"
	}
	else
	{
		if (inObj.validrule_id == "")
		{
		document.all.VALID_ID.value = "null"
		document.all.VALID_TEXT.innerHTML = ""
		}
		else
		{
		document.all.VALID_ID.value = inObj.validrule_id
		document.all.VALID_TEXT.innerHTML = (inObj.validrule_id == "null") ? "" : inObj.validrule_id
		}
	}
	
	if (inObj.persistrule_id == "-999999999")
	{
		document.all.PERSIST_ID.value = "-999999999"
		document.all.PERSIST_TEXT.innerHTML = "*Using attribute defined value*"
		document.all.PERSIST_TEXT.title = "*Using attribute defined value*"
		document.all.PERSIST_TEXT.style.color = "Maroon"
	}
	else
	{
		document.all.PERSIST_ID.value = inObj.persistrule_id
		document.all.PERSIST_TEXT.innerHTML = (inObj.persistrule_id == "null") ? "" : inObj.persistrule_id
	}

	if (inObj.action_id == "-999999999")
	{
		document.all.ACTION_ID.value = "-999999999"
		document.all.ACTION_TEXT.innerHTML = "*Using attribute defined value*"
		document.all.ACTION_TEXT.title = "*Using attribute defined value*"
		document.all.ACTION_TEXT.style.color = "Maroon"
	}
	else
	{
		document.all.ACTION_ID.value = inObj.action_id
		document.all.ACTION_TEXT.innerHTML = (inObj.action_id == "null") ? "" : inObj.action_id
	}

	
	if (inObj.lu_type_id == "-999999999")
	{
		document.all.LU_TYPE_ID.value = "-999999999"
		document.all.LOOKUPNAME_TEXT.innerHTML = "*Using attribute defined value*"
		document.all.LOOKUPNAME_TEXT.style.color = "Maroon"
	}
	else
	{
		document.all.LU_TYPE_ID.value = inObj.lu_type_id
		document.all.LOOKUPNAME_TEXT.innerHTML = (inObj.lu_type_id == "null") ? "" : inObj.lu_type_id
	}
	
	if (inObj.spellcheck_flg =="U") 
	{
	UNKNOWNSPELLCHECK.innerHTML = "*"
	UNKNOWNSPELLCHECK.style.color = "maroon"
	}
	else
	{
		if (inObj.spellcheck_flg=="Y") 
		{spellcheck_flg.checked = true;}
	}
	
	

	helpstring.value = inObj.helpstring
	if ( inObj.helpstring == "-999999999")
	{helpstring.style.color = "Maroon"}
	
	description.value = inObj.description
	if (inObj.description == "-999999999")
	{description.style.color = "Maroon"	}
	

		if (inObj.mandatory == true)
		{
		PropMANDATORY_FLG.checked = true
		}
						
		if ( inObj.ludisplay_flg == true)
		{
		LUDISPLAY_FLG.checked = true
		}
			
		if (inObj.lustorage_flg == true)
		{
		LUSTORAGE_FLG.checked = true
		}
        if (inObj.reapplyoverride_flg == "Y") 
		{reapplyoverride_flg.checked = true;}
		
		inObj.pagestatus = "cancel";
}
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function BtnOK_onclick() 
{
		inObj.SampleValue = PropSAMPLEVALUE.value;
		//Obj.label = PropLABEL.value;
		inObj.xpos = PropXPOS.value; 
		inObj.ypos = PropYPOS.value;
		inObj.width = PropWIDTH.value;
		inObj.height = PropHEIGHT.value;
		inObj.name = PropNAME.value;
		inObj.SampleValue = PropSAMPLEVALUE.value;
		inObj.sequence = PropSEQUENCE.value;
		inObj.lucolumn_name = LUCOLUMN_NAME.value;
		inObj.lustorage_name = LUSTORAGE_NAME.value;
		inObj.mandatory = PropMANDATORY_FLG.checked;
		
		inObj.ludisplay_flg = LUDISPLAY_FLG.checked;
		inObj.lustorage_flg = LUSTORAGE_FLG.checked;
		
		inObj.type = PropTYPE.value;
		inObj.pagestatus = "save";		
//New		
		inObj.defaultvalue = defaultvalue.value
		inObj.unknownvalue = unknownvalue.value
		inObj.textlength = textlength.value		
		inObj.helpstring = helpstring.value
		inObj.description = description.value
		inObj.caption = PropCaption.value
		inObj.inputtype = inputtype.value
		inObj.entrymask = entrymask.value

		if (UNKNOWNVALIDVALUE.innerHTML == "*")
		{ inObj.validvaluefield_flg = 'U' }
		else
		{
			if (validvaluefield_flg.checked == true)
				{inObj.validvaluefield_flg = 'Y' }
			else
				{inObj.validvaluefield_flg = 'N' }
		}	
		
		if (UNKNOWNSPELLCHECK.innerHTML == "*")
		{inObj.spellcheck_flg = 'U'}
		else
		{
			if (spellcheck_flg.checked == true)
				{inObj.spellcheck_flg = 'Y' }
			else
				{inObj.spellcheck_flg = 'N' }
		}
			
		if (reapplyoverride_flg.checked == true)
				{inObj.reapplyoverride_flg = 'Y' }
		else
				{inObj.reapplyoverride_flg = 'N' }
		

		if (	LU_TYPE_ID.value != "")
		{inObj.lu_type_id = LU_TYPE_ID.value}
		else
		{inObj.lu_type_id = "null"}
			
			
		if (	VISIBLE_ID.value != "")
		{inObj.visiblerule_id = VISIBLE_ID.value}
		else
		{inObj.visiblerule_id = "null"}
		
		if (ENABLED_ID.value != "")
		{inObj.enabledrule_id = ENABLED_ID.value}
		else
		{inObj.enabledrule_id = "null"}

		if (ATTRIBUTEFRAME_ID.value != "")
		{inObj.attributeframe_id = ATTRIBUTEFRAME_ID.value}
		else
		{inObj.attributeframe_id = "null"}
		
		if (VALID_ID.value != "") 
		{inObj.validrule_id = VALID_ID.value}
		else
		{inObj.validrule_id = "null"}
		
		if (PERSIST_ID.value != "")
		{inObj.persistrule_id = PERSIST_ID.value}
		else
		{inObj.persistrule_id = "null"}
		
		if (ACTION_ID.value != "")
		{inObj.action_id = ACTION_ID.value}
		else
		{inObj.action_id = "null"}
		
lret = PageCheck()
	if (true == lret) 			
		{
			window.returnvalue = inObj;
			window.close();
		}
		
}

function BtnCancel_onclick() {
inObj.pagestatus = "cancel";
window.returnvalue = null;
window.close();
}

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
function CFrameSearchObj()
{
	this.FrameID = "";
}
var FrameObj = new CFrameSearchObj();
var RuleSearchObj = new CRuleSearchObj();
var g_StatusInfoAvailable = false;
var LookupTypeSearchObj = new CLookupTypeSearchObj();
//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

<!--#include file="..\lib\Help.asp"-->

Sub document_onkeydown
Select Case window.event.keyCode
	Case 13
		Call BtnOK_onclick()
	Case Else
End Select
End Sub

Function DetachRule(ID, SPANID)
	ID.value = ""
	SPANID.innerText = ""
End Function

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
	'MODE = document.body.getAttribute("ScreenMode")
	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false
	If (RID = "") Or (Not IsNumeric(RID)) Then RID = "NEW"
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID &  "&TITLE=" & strTITLE & "&MODE=" & MODE 
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = RuleSearchObj.RID
		end if
		UpdateRuleText(SPANID)
	ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		'UpdateRuleText(SPANID)
	End If
End Function

Function SetUnknown(ID, SPANID, ELTYPE, flgnm)
Select Case ELTYPE
	Case "RULE"
		ID.value = "-999999999"
		SPANID.innerHtml = "*Using attribute defined value*"
		SPANID.Title = "*Using attribute defined value*"
		SPANID.style.color = "Maroon"
	Case "FLAG"
		Select Case flgnm
			Case "validvaluefield_flg"
				inObj.validvaluefield_flg = "U"
				UNKNOWNVALIDVALUE.innerHTML = "*"
				UNKNOWNVALIDVALUE.style.color = "Maroon"
			Case "spellcheck_flg"
				inObj.spellcheck_flg = "U"
				UNKNOWNSPELLCHECK.innerHTML = "*"
				UNKNOWNSPELLCHECK.style.color = "Maroon"
		End Select
	Case "DROPDOWN"
		ID.Value = "-999999999"
		ID(ID.selectedindex).Style.color = "Maroon"
	Case "TEXT"
		ID.Value = "-999999999"
		ID.Style.color = "Maroon"
End Select
End Function

Sub validvaluefield_flg_onclick
	UNKNOWNVALIDVALUE.innerHTML = ""
End Sub

Sub spellcheck_flg_onclick
	UNKNOWNSPELLCHECK.innerHTML = ""
End Sub

Sub PropCaption_onkeypress
	PropCaption.style.color = "Black"
End Sub

Sub helpstring_onkeypress
	helpstring.style.color = "black"
End Sub

Sub entrymask_onkeypress
	entrymask.style.color = "Black"
End Sub

Sub description_onkeypress
	description.style.color = "Black"
End Sub

Sub defaultvalue_onkeypress
	defaultvalue.style.color = "Black"
End Sub

Sub textlength_onkeypress
	textlength.style.color = "Black"
End Sub

Sub unknownvalue_onkeypress
	unknownvalue.style.color = "Black"
End Sub

Sub BtnAttachLookupType_OnClick
	LUTID = document.all.LU_TYPE_ID.value
	MODE = document.body.getAttribute("ScreenMode")

	LookupTypeSearchObj.LUTID = LUTID
	LookupTypeSearchObj.LUTIDName = document.all.LOOKUPNAME_TEXT.innerText
	LookupTypeSearchObj.Selected = false

	If LUTID = "" Then LUTID = "NEW"
	
	If LUTID = "NEW" And MODE = "RO" Then
		MsgBox "No lookup type currently attached.",0,"FNSNetDesigner"
		Exit Sub
	End If
	
	strURL = "..\LookupType\LookupTypeMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&LUTID=" & LUTID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,LookupTypeSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If LookupTypeSearchObj.Selected = true Then
		If LookupTypeSearchObj.LUTID <> document.all.LU_TYPE_ID.value then
			document.all.LU_TYPE_ID.value = LookupTypeSearchObj.LUTID
			document.all.LOOKUPNAME_TEXT.style.color = "Black"
		end if
		UpdateLookupText (document.all.LOOKUPNAME_TEXT)
	ElseIf document.all.LU_TYPE_ID.value = LookupTypeSearchObj.LUTID And LookupTypeSearchObj.LUTID <> "" Then
		UpdateLookupText (document.all.LOOKUPNAME_TEXT)
	End If
End Sub

Sub UpdateLookupText (SPANID)
		SPANID.innertext = LookupTypeSearchObj.LUTID
		SPANID.title = LookupTypeSearchObj.LUTIDName
End Sub


Sub BtnDetachFrame_onclick
	document.all.ATTRIBUTEFRAME_ID.value = ""
	document.all.ATTRIBUTEFRAME_TEXT.innerHTML = ""
End Sub

Sub BtnAttachFrame_onclick
	strURL = "FrameSearchModal.asp"
	showModalDialog  strURL  ,FrameObj ,"dialogWidth:450px;dialogHeight:450px;center"
	If FrameObj.FrameID <> "" Then
		document.all.ATTRIBUTEFRAME_ID.value = FrameObj.FrameID
		document.all.ATTRIBUTEFRAME_TEXT.innerHTML = FrameObj.FrameID
		document.all.ATTRIBUTEFRAME_TEXT.Style.Color = "Black"
	End If
End Sub
-->
</SCRIPT>
</HEAD>
<BODY  bGCOLOR="#d6cfbd">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Instance Properties</SPAN>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL COLSPAN=2>Name/Attribute:<BR>
<INPUT TYPE=TEXT CLASS=DISABLED READONLY NAME=PropNAME SIZE=80 MAXLENGTH=80></TD>
</TR>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  ID=UNKNOWN OnClick='SetUnknown PropCaption,PropCaption, "TEXT", ""'>
Caption:<BR>
<INPUT TYPE=TEXT CLASS=LABEL ID="PropCaption" NAME="PropCaption" SIZE=40 MAXLENGTH=80></TD>
<TD CLASS=LABEL >Sample Value:<BR>
<INPUT TYPE=TEXT CLASS=LABEL ID="PropSAMPLEVALUE" NAME="PropSAMPLEVALUE" SIZE=40 MAXLENGTH=80></TD>
</TR>
</TABLE>
<TABLE CELLPADDING=3 CELLSPACING=0>
<TR>
<TD CLASS=LABEL>XPos:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL NAME="PropXPOS" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>YPos:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL NAME="PropYPOS" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>Width:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL NAME="PropWIDTH" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>Height:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL NAME="PropHEIGHT" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>Sequence:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL NAME="PropSEQUENCE" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>
Type:<BR>
<SELECT NAME=PropTYPE CLASS=LABEL>
<OPTION VALUE="DATA ENTRY">Data Entry
<OPTION VALUE="LOOK UP">Look Up
</SELECT>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>LU Column Name:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=LUCOLUMN_NAME SIZE=40></TD>
<TD CLASS=LABEL VALIGN=MIDDLE>LU Storage Name:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=LUSTORAGE_NAME SIZE=40></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  ID=UNKNOWN OnClick='SetUnknown inputtype,inputtype, "DROPDOWN", ""'>
Input Type:<br><SELECT NAME="inputtype" CLASS="LABEL">
<%=GetValidValuesHTML("ATTRIBUTE_INPUTTYPE",RSINPUTTYPE,true)%>
<OPTION VALUE="-999999999">-999999999
</SELECT>
</TD>
<TD CLASS=LABEL VALIGN=MIDDLE><INPUT TYPE=CHECKBOX NAME=LUDISPLAY_FLG CLASS=LABEL>
LU Display Flag?</TD>
<TD CLASS=LABEL><INPUT TYPE=CHECKBOX NAME="LUSTORAGE_FLG" CLASS=LABEL>
LU Storage Flag?</TD>
</TR>
<TR>
<TD CLASS=LABEL>
<INPUT TYPE=CHECKBOX NAME=PropMANDATORY_FLG CLASS=LABEL>
Mandatory?</TD>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown validvaluefield_flg,validvaluefield_flg, "FLAG", "validvaluefield_flg"'>
<INPUT TYPE=CHECKBOX NAME=validvaluefield_flg CLASS=LABEL>
Valid Value?<SPAN ID=UNKNOWNVALIDVALUE CLASS=LABEL></SPAN></TD>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown spellcheck_flg,spellcheck_flg, "FLAG", "spellcheck_flg"'>
<INPUT TYPE=CHECKBOX NAME=spellcheck_flg CLASS=LABEL>
Spell Check?<SPAN ID=UNKNOWNSPELLCHECK CLASS=LABEL></SPAN></TD>
<TD CLASS=LABEL><INPUT TYPE=CHECKBOX NAME=reapplyoverride_flg CLASS=LABEL >
Reapply Ovrd?</TD>
</TR>
</TABLE>
<table class="LABEL">
		<td nowrap>
		<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown LU_TYPE_ID,LOOKUPNAME_TEXT, "RULE", ""'>
		<IMG NAME=BtnAttachLookupType STYLE="CURSOR:HAND" TITLE="Attach Lookup Type" SRC="..\IMAGES\Attach.gif">
		<IMG NAME=BtnDetachLookupType STYLE="CURSOR:HAND" TITLE="Detach Lookup Type" OnClick="VBScript::DetachRule LU_TYPE_ID, LOOKUPNAME_TEXT" SRC="..\IMAGES\Detach.gif"></td>
		<td nowrap>Lookup ID:&nbsp<SPAN ID=LOOKUPNAME_TEXT CLASS=LABEL TITLE="" ></SPAN>
		<input type="hidden" name="LU_TYPE_ID"></td>
	<td>	
	<td>
</table>	
<table class="LABEL">
		<td nowrap>
		<IMG NAME=BtnAttachFrame STYLE="CURSOR:HAND" TITLE="Attach Frame" SRC="..\IMAGES\Attach.gif">
		<IMG NAME=BtnDetachFrame STYLE="CURSOR:HAND" TITLE="Detach Frame" OnClick="VBScript::DetachRule ATTRIBUTEFRAME_ID, ATTRIBUTEFRAME_TEXT" SRC="..\IMAGES\Detach.gif"></td>
		<td nowrap>AttributeFrame ID:&nbsp<SPAN ID=ATTRIBUTEFRAME_TEXT CLASS=LABEL TITLE="" ></SPAN>
		<input type="hidden" name="ATTRIBUTEFRAME_ID"></td>
	<td>	
	<td>
</table>	
<TABLE>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown entrymask,entrymask, "TEXT", ""'>
Entry Mask:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=entrymask MAXLENGTH=80 SIZE=40></TD>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown defaultvalue,defaultvalue, "TEXT", ""'>
Default Value:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=defaultvalue MAXLENGTH=255 SIZE=40></TD>
</TR>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown unknownvalue,unknownvalue, "TEXT", ""'>
Unknown Value:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=unknownvalue MAXLENGTH=255 SIZE=40></TD>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif"  STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWNHELPSTRING OnClick='SetUnknown textlength,textlength, "TEXT", ""'>
Text Length:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=textlength MAXLENGTH=255 SIZE=40></TD>
</TR>
<TR>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown helpstring,helpstring, "TEXT", ""'>
Help String:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=helpstring MAXLENGTH=255 SIZE=40></TD>
<TD CLASS=LABEL><IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWNDESCRIPTION OnClick='SetUnknown description,description, "TEXT", ""'>
Description:<BR><INPUT CLASS=LABEL TYPE=TEXT NAME=description MAXLENGTH=255 SIZE=40></TD>
</TR>
</TABLE>


<TABLE CLASS=LABEL>
<tr>
<td>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown VISIBLE_ID,VISIBLE_TEXT, "RULE", ""'>
<IMG NAME=BtnAttachVisible STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule VISIBLE_ID, VISIBLE_TEXT,'Visible'">
<IMG NAME=BtnDetachVisible STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule VISIBLE_ID, VISIBLE_TEXT">
</TD>
<td nowrap>Visible Rule:</td>
<td><SPAN ID=VISIBLE_TEXT CLASS=LABEL TITLE=""></SPAN><INPUT TYPE=HIDDEN NAME=VISIBLE_ID VALUE="<%= RSVISIBLERULE_ID %>"></TD>
</tr>
<tr>
<td>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown ENABLED_ID,ENABLED_TEXT, "RULE", ""'>
<IMG NAME=BtnAttachEnabled STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule ENABLED_ID, ENABLED_TEXT,'Enabled'">
<IMG NAME=BtnDetachEnabled STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule ENABLED_ID, ENABLED_TEXT">
</td>
<td nowrap>Enabled Rule:</td>
<td><SPAN ID=ENABLED_TEXT CLASS=LABEL TITLE=""></SPAN><INPUT TYPE=HIDDEN NAME=ENABLED_ID VALUE="<%= RSENABLEDRULE_ID %>"></TD>
</tr>
<tr>
<td>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown VALID_ID,VALID_TEXT, "RULE", ""'>
<IMG NAME=BtnAttachValid STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule VALID_ID, VALID_TEXT,'Valid'">
<IMG NAME=BtnDetachValid STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule VALID_ID, VALID_TEXT">
</td>
<td nowrap>Valid Rule:</td>
<td><SPAN ID=VALID_TEXT CLASS=LABEL TITLE=""></SPAN><INPUT TYPE=HIDDEN NAME=VALID_ID VALUE="<%= RSVALIDRULE_ID %>"></TD>
</tr>
<td>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown PERSIST_ID,PERSIST_TEXT, "RULE", ""'>
<IMG NAME=BtnAttachPersist STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule PERSIST_ID, PERSIST_TEXT,'Persist'">
<IMG NAME=BtnDetachPersist STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule PERSIST_ID, PERSIST_TEXT">
</TD>
<td nowrap>Persist Rule:</td>
<td><SPAN ID=PERSIST_TEXT CLASS=LABEL TITLE=""></SPAN><INPUT TYPE=HIDDEN NAME=PERSIST_ID VALUE="<%= RSPERSISTRULE_ID %>"></TD>
</tr>
<td>
<IMG SRC="../Images/UnknownValue.gif" STYLE="CURSOR:HAND"  TITLE="Use attribute defined value" ID=UNKNOWN OnClick='SetUnknown ACTION_ID,ACTION_TEXT, "RULE", ""'>
<IMG NAME=BtnAttachAction STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule ACTION_ID, ACTION_TEXT,'Action'">
<IMG NAME=BtnDetachAction STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::DetachRule ACTION_ID, ACTION_TEXT">
<td nowrap>Action:</td>
<td><SPAN ID=ACTION_TEXT CLASS=LABEL TITLE=""></SPAN><INPUT TYPE=HIDDEN NAME=ACTION_ID VALUE="<%= RSACTION_ID %>"></TD>
</tr>
</table>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnOK LANGUAGE=javascript onclick="return BtnOK_onclick()">Ok</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel LANGUAGE=javascript onclick="return BtnCancel_onclick()">Cancel</BUTTON></TD>
</TR>
</TABLE>
</BODY>
</HTML>
