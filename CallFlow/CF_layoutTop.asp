<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then   
	Session("NAME") = ""
	Response.Redirect "CF_LayoutTop.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=JavaScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
<!--#include file="..\lib\Help.asp"-->

Dim Zoom_Toggle, Dirty_Flag,Xinc,Yinc
Xinc = 3
Yinc = 3
Dirty_Flag = 0
Zoom_Toggle = 0

Function ConSingle( Val )
	Select Case Val
		Case "true"
			ConSingle = "Y"
		Case "True"
			ConSingle = "Y"
		Case "TRUE"
			ConSingle = "Y"
		Case "false"
			ConSingle = "N"
		Case "False"
			ConSingle = "N"
		Case "FALSE"	
			ConSingle = "N"
		Case Else
			ConSingle = "N"
	End Select
End Function

Sub BtnDelete_onclick
<% If MODE = "RO" Then Response.write(" Exit Sub ") %>
	Set X = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem
	If IsObject(X) Then
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.DeleteItem(x.RefID)
		ILength = ListAttributes.options.length
		If ListAttributes.options.length > 0 Then
			For i = 0 to ILength-1  Step 1
				If x.RefID = ListAttributes(i).value Then
					ListAttributes.Remove(i)
					Exit Sub
				End If
			Next
		End if
	End If
End Sub

Sub BtnZoom_onclick
	If Zoom_Toggle = 0 Then
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.ZoomIn = True
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.WIDTH=2000
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.HEIGHT=2080
		Parent.Frames("LAYOUTAREA").document.all.HEADER.style.width = "204%"
		Parent.Frames("LAYOUTAREA").document.all.HEADER.style.fontsize = "24pt"
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.style.top = "54px"
		ZoomMenuText = "Zoom Out"
		Zoom_Toggle = 1
	Else
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.ZoomIn = False
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.WIDTH=1000
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.HEIGHT=1040
		Parent.Frames("LAYOUTAREA").document.all.HEADER.style.width = "102%"
		Parent.Frames("LAYOUTAREA").document.all.HEADER.style.fontsize = "12pt"
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.style.top = "27px"
		ZoomMenuText = "Zoom In"
		Zoom_Toggle = 0
	End If
	Parent.Frames("LAYOUTAREA").SetMenu(ZoomMenuText)
End Sub

Sub UNDO_onclick
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.Undo
End Sub

Sub ListAttributes_OnChange
	Parent.Frames("LAYOUTAREA").LayoutCtl.SetSelectedItem(ListAttributes.Value)
	Set X = Parent.Frames("LAYOUTAREA").LayoutCtl.GetSelectedItem
	If IsObject(X) Then

		If x.xpos < 0 or x.ypos < 0 Then
			call BtnProperties_onclick
		Else
			XPIX = Parent.Frames("LAYOUTAREA").LayoutCtl.GetXPixelPos(x.xpos) 
			YPIX = Parent.Frames("LAYOUTAREA").LayoutCtl.GetYPixelPos(x.ypos) 
			Parent.Frames("LAYOUTAREA").window.scrollto XPIX, YPIX
		End If
	End If
End Sub

Sub BtnSave_onclick
<% If MODE = "RO" Then Response.write(" Exit Sub ") %>
sStr = ""
InsStr = ""
delResult = ""
UpCount = 0
InCount = 0
delCount = 0


	Set objCol = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.PageItems
	For Each x In objCol
		If x.dirty = "True" AND x.Status = "MODIFY" AND x.deleted = false Then
			UpCount = UpCount + 1
			sStr = sStr & "ATTR_INSTANCE_ID" & Chr(129) &  x.refid & chr(129) & "0" & chr(128)
		    sStr = sStr & "XPOS" & Chr(129) & x.XPos & chr(129) & "0" & chr(128)
			sStr = sStr & "YPOS" & Chr(129) & x.YPos & chr(129) & "0" & chr(128)
			sStr = sStr & "WIDTH" & Chr(129) & x.Width & chr(129) & "0" & chr(128)
			sStr = sStr & "HEIGHT" & Chr(129) & x.Height & chr(129) & "0" & chr(128)
			sStr = sStr & "TYPE" & Chr(129) & x.GetExtraProperty("TYPE")  & chr(129) & "1" & chr(128)
			sStr = sStr & "MANDATORY_FLG" & Chr(129) & ConSingle(x.mandatoryflag) & chr(129) & "1" & chr(128)
			sStr = sStr & "LUCOLUMN_NAME" & Chr(129) & x.GetExtraProperty("LUCOLUMN_NAME") & chr(129) & "1" & chr(128)
			sStr = sStr & "LUDISPLAY_FLG" & Chr(129) & ConSingle(x.GetExtraProperty("LUDISPLAY_FLG")) & chr(129) & "1" & chr(128)
			sStr = sStr & "LUSTORAGE_FLG" & Chr(129) & ConSingle(x.GetExtraProperty("LUSTORAGE_FLG")) & chr(129) & "1" & chr(128)
			sStr = sStr & "LUSTORAGE_NAME" & Chr(129) & x.GetExtraProperty("LUSTORAGE_NAME") & chr(129) & "1" & chr(128)
			sStr = sStr & "SEQUENCE" & Chr(129) & x.sequence & chr(129) & "0" & chr(128)

			If x.GetExtraProperty("LU_TYPE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("LU_TYPE_ID")
			End If			
			sStr = sStr & "LU_TYPE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			sStr = sStr & "CAPTION" & Chr(129) & x.GetExtraProperty("CAPTION") & chr(129) & "1" & chr(128)
			sStr = sStr & "INPUTTYPE" & Chr(129) & x.GetExtraProperty("INPUTTYPE") & chr(129) & "1" & chr(128)
			sStr = sStr & "ENTRYMASK" & Chr(129) & x.GetExtraProperty("ENTRYMASK") & chr(129) & "1" & chr(128)
			sStr = sStr & "VALIDVALUEFIELD_FLG" & Chr(129) & x.GetExtraProperty("VALIDVALUEFIELD_FLG") & chr(129) & "1" & chr(128)
			sStr = sStr & "DEFAULTVALUE" & Chr(129) & x.GetExtraProperty("DEFAULTVALUE") & chr(129) & "1" & chr(128)
			sStr = sStr & "UNKNOWNVALUE" & Chr(129) & x.GetExtraProperty("UNKNOWNVALUE") & chr(129) & "1" & chr(128)
			sStr = sStr & "TEXTLENGTH" & Chr(129) & x.GetExtraProperty("TEXTLENGTH") & chr(129) & "1" & chr(128)

			If x.GetExtraProperty("VISIBLERULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("VISIBLERULE_ID")
			End If
			sStr = sStr & "VISIBLERULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			
			
			If x.GetExtraProperty("ENABLEDRULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("ENABLEDRULE_ID")
			End If
			sStr = sStr & "ENABLEDRULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			
	
			If x.GetExtraProperty("VALIDRULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("VALIDRULE_ID")
			End If		
			sStr = sStr & "VALIDRULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)


			If x.GetExtraProperty("PERSISTRULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("PERSISTRULE_ID")
			End If	
			sStr = sStr & "PERSISTRULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)


			If x.GetExtraProperty("ACTION_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("ACTION_ID")
			End If	
			sStr = sStr & "ACTION_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)

			sStr = sStr & "SPELLCHECK_FLG" & Chr(129) & x.GetExtraProperty("SPELLCHECK_FLG") & chr(129) & "1" & chr(128)
			sStr = sStr & "REAPPLYOVERRIDE_FLG" & Chr(129) & x.GetExtraProperty("REAPPLYOVERRIDE_FLG") & chr(129) & "1" & chr(128)
			sStr = sStr & "HELPSTRING" & Chr(129) & x.GetExtraProperty("HELPSTRING") & chr(129) & "1" & chr(128)
	
			If x.GetExtraProperty("ATTRIBUTEFRAME_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("ATTRIBUTEFRAME_ID")
			End If	
			sStr = sStr & "ATTRIBUTEFRAME_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
	
			sStr = sStr & "DESCRIPTION" & Chr(129) & x.GetExtraProperty("DESCRIPTION") & chr(129) & "1" & chr(128)
			sStr = sStr & chr(130)
		End If

		If X.Dirty AND x.Status = "NEW" Then
			InCount = InCount + 1
			InsStr = InsStr & "ATTRIBUTE_ID" & Chr(129) &  x.GetExtraProperty("ATTRIBUTE_ID") & chr(129) & "0" & chr(128)
			InsStr = InsStr & "FRAME_ID" & Chr(129) & "<%= Trim(Request.QueryString("FRAMEID")) %>" & Chr(129) & "0" & chr(128)
			InsStr = InsStr & "ATTR_INSTANCE_ID" & Chr(129)  &  x.refid & chr(129) & "0" & chr(128)
		    InsStr = InsStr & "XPOS" & Chr(129) & x.XPos & chr(129) & "0" & chr(128)
			InsStr = InsStr & "YPOS" & Chr(129) & x.YPos & chr(129) & "0" & chr(128)
			InsStr = InsStr & "WIDTH" & Chr(129) & x.Width & chr(129) & "0" & chr(128)
			InsStr = InsStr & "HEIGHT" & Chr(129) & x.Height & chr(129) & "0" & chr(128)
			InsStr = InsStr & "TYPE" & Chr(129) & x.GetExtraProperty("TYPE")  & chr(129) & "1" & chr(128)
			InsStr = InsStr & "MANDATORY_FLG" & Chr(129) & ConSingle(x.mandatoryflag) & chr(129) & "1" & chr(128)
			InsStr = InsStr & "LUCOLUMN_NAME" & Chr(129) & x.GetExtraProperty("LUCOLUMN_NAME") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "LUDISPLAY_FLG" & Chr(129) & ConSingle(x.GetExtraProperty("LUDISPLAY_FLG")) & chr(129) & "1" & chr(128)
			InsStr = InsStr & "LUSTORAGE_FLG" & Chr(129) & ConSingle(x.GetExtraProperty("LUSTORAGE_FLG")) & chr(129) & "1" & chr(128)
			InsStr = InsStr & "LUSTORAGE_NAME" & Chr(129) & x.GetExtraProperty("LUSTORAGE_NAME") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "SEQUENCE" & Chr(129) & x.sequence & Chr(129) & "0" & chr(128)
	
			If x.GetExtraProperty("LU_TYPE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("LU_TYPE_ID")
			End If
			InsStr = InsStr & "LU_TYPE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			InsStr = InsStr & "CAPTION" & Chr(129) & x.GetExtraProperty("CAPTION") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "INPUTTYPE" & Chr(129) & x.GetExtraProperty("INPUTTYPE") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "ENTRYMASK" & Chr(129) & x.GetExtraProperty("ENTRYMASK") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "VALIDVALUEFIELD_FLG" & Chr(129) & x.GetExtraProperty("VALIDVALUEFIELD_FLG") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "DEFAULTVALUE" & Chr(129) & x.GetExtraProperty("DEFAULTVALUE") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "UNKNOWNVALUE" & Chr(129) & x.GetExtraProperty("UNKNOWNVALUE") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "TEXTLENGTH" & Chr(129) & x.GetExtraProperty("TEXTLENGTH") & chr(129) & "1" & chr(128)

			If x.GetExtraProperty("VISIBLERULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("VISIBLERULE_ID")
			End If
			InsStr = InsStr & "VISIBLERULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			
			
			If x.GetExtraProperty("ENABLEDRULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("ENABLEDRULE_ID")
			End If
			InsStr = InsStr & "ENABLEDRULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			
	
			If x.GetExtraProperty("VALIDRULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("VALIDRULE_ID")
			End If		
			InsStr = InsStr & "VALIDRULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)


			If x.GetExtraProperty("PERSISTRULE_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("PERSISTRULE_ID")
			End If	
			InsStr = InsStr & "PERSISTRULE_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)


			If x.GetExtraProperty("ACTION_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("ACTION_ID")
			End If	
			InsStr = InsStr & "ACTION_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)
			InsStr = InsStr & "SPELLCHECK_FLG" & Chr(129) & x.GetExtraProperty("SPELLCHECK_FLG") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "REAPPLYOVERRIDE_FLG" & Chr(129) & x.GetExtraProperty("REAPPLYOVERRIDE_FLG") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "HELPSTRING" & Chr(129) & x.GetExtraProperty("HELPSTRING") & chr(129) & "1" & chr(128)
			InsStr = InsStr & "DESCRIPTION" & Chr(129) & x.GetExtraProperty("DESCRIPTION") & chr(129) & "1" & chr(128)

			If x.GetExtraProperty("ATTRIBUTEFRAME_ID") = "" Then
				sVal =  "null"	
			Else
				sVal = x.GetExtraProperty("ATTRIBUTEFRAME_ID")
			End If	
			InsStr = InsStr & "ATTRIBUTEFRAME_ID" & Chr(129) & sVal & chr(129) & "0" & chr(128)

			InsStr = InsStr & chr(130)
			x.status = "MODIFY"
		End if
		If x.deleted = True and x.status = "DELETED" Then
			x.status = "DELETED"
			delCount = delCount + 1
			delResult = delResult & x.refid  & chr(130)
		End if	
	Next
	

	If UpCount > 0 OR InCount > 0 OR DelCount > 0 Then
		document.all.TxtUpdateData.Value = sStr
		document.all.TxtInsertData.Value = InsStr
		document.all.TxtDeleteData.Value = delResult
		document.all.UpCOUNT.Value = UpCount
		document.all.InCOUNT.Value = InCount
		document.all.DELCOUNT.Value = DelCount
		
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.CleanAllDirty
		document.all.SAVEDATA.Submit()
		StatusSpan.innerHTML = "Saved Successfully"
		StatusSpan.style.color = "MAROON"
		If InCount > 0 Then
			Parent.frames("LAYOUTAREA").location.href = "CF_LayoutBottom.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>"
		End If
	Else
		StatusSpan.innerHTML = "Nothing to Save"
		StatusSpan.style.color = "MAROON"
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.CleanAllDirty
	End If
	
End Sub

Function BtnAddAttribute_onclick()
<% If MODE = "RO" Then Response.write(" Exit Function ") %>
AttributeSearchObj.Selected = false

	showModalDialog  "../Attribute/AttributeMaintenance.asp"  , AttributeSearchObj ,"dialogWidth=520px; dialogHeight=580px; center=yes"
If AttributeSearchObj.Selected <> false Then
	listarray = split(AttributeSearchObj.AIDName, "||")
	IDarray = split (AttributeSearchObj.AID, "||")

	If AttributeSearchObj.AIDCaption <> "" Then
		CapArray = split(AttributeSearchObj.AIDCaption, "||")
	Else
		Dim CapArray(1)
		CapArray(0) = ""
	End If
	For i = 0 to Ubound(listarray) step 1
		SequenceNumber = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.NextSequenceNumber
		Set objCol = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.PageItems
		Set NewObj = objCol.AddItem ("NEW" & SequenceNumber, listarray(i),CapArray(i)  , Xinc,Yinc, 15, 4, "Courier New", 9,false ,false ,false ,false ,false , SequenceNumber, "Output Field", "NEW",  "", False	)
		NewObj.MandatoryFlag = False
		
		NewObj.SetExtraProperty "LUCOLUMN_NAME", true, ""
		NewObj.SetExtraProperty "LUDISPLAY_FLG", true, false
		NewObj.SetExtraProperty "LUSTORAGE_FLG", true, false
		NewObj.SetExtraProperty "LUSTORAGE_NAME", true, ""
		NewObj.SetExtraProperty "ATTRIBUTE_ID", true, IDarray(i)
		NewObj.SetExtraProperty "TYPE" , true, "DATA ENTRY"
		
	NewObj.SetExtraProperty "ATTR_CAP", true, CapArray(i)
	NewObj.SetExtraProperty "CAPTION", true, "-999999999"
	NewObj.SetExtraProperty "INPUTTYPE", true, "-999999999"
	NewObj.SetExtraProperty "ENTRYMASK", true, "-999999999"
	NewObj.SetExtraProperty "VALIDVALUEFIELD_FLG", true, "U"
	NewObj.SetExtraProperty "DEFAULTVALUE", true ,"-999999999"
	NewObj.SetExtraProperty "UNKNOWNVALUE", true, "-999999999"
	NewObj.SetExtraProperty "TEXTLENGTH", true, "-999999999"
	NewObj.SetExtraProperty "VISIBLERULE_ID", true, "-999999999"
	NewObj.SetExtraProperty "ENABLEDRULE_ID", true, "-999999999"
	NewObj.SetExtraProperty "VALIDRULE_ID", true, "-999999999"
	NewObj.SetExtraProperty "PERSISTRULE_ID", true, "-999999999"
	NewObj.SetExtraProperty "ACTION_ID", true, "-999999999"
	NewObj.SetExtraProperty "SPELLCHECK_FLG", true, "U"
	NewObj.SetExtraProperty "REAPPLYOVERRIDE_FLG", true, "N"
	NewObj.SetExtraProperty "HELPSTRING", true, "-999999999"
	NewObj.SetExtraProperty "DESCRIPTION", true, "-999999999"
	NewObj.SetExtraProperty "LU_TYPE_ID", true, "-999999999"
	NewObj.SetExtraProperty "ATTRIBUTEFRAME_ID", true, "null"
	
					
		Xinc = Xinc  +2
		Yinc = Yinc + 2
		Set objOption = document.createElement("option")
		objOption.value = SequenceNumber
		objOption.Text = listarray(i)
		ListAttributes.add( objOption )
	Next
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.SetSelectedItem("NEW" & SequenceNumber)
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems
End If
End Function

-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function CItemClass()
{
	this.SampleValue = "One";
	this.name = "Two";
	this.attributekey = "Three";
	this.formatstring = "Four";
	this.label = "five";
	this.sequence = "six";
	this.type = "seven";
	this.mandatory = "eight";
	this.lucolumn_name = "nine";
	this.ludisplay_flg = "fourteen";
	this.lustorage_flg = "fifteen";
	this.lustorage_name = "sixteen";
	this.xpos = "ten";
	this.ypos = "eleven";
	this.width = "twelve";
	this.height = "thirteen";
	this.pagestatus = "cancel";
	this.attribute_id = "ID";
	
	this.attributeframe_id = "";
	this.caption = "";
	this.inputtype = "";
	this.entrymask = "";
	this.validvaluefield_flg = "";
	this.defaultvalue = "";
	this.unknownvalue = "";
	this.textlength = "";
	this.visiblerule_id = "";
	this.enabledrule_id = "";
	this.validrule_id = "";
	this.persistrule_id = "";
	this.action_id = "";
	this.spellcheck_flg = "";
	this.reapplyoverride_flg = "";
	this.helpstring = "";
	this.description = "";
	this.lu_type_id = "";
	
			
}
function CMyOptions()
{
	this.defaultfont = "One";
	this.defaultfontsize = "Two";
	this.defaultheight = "three";
	this.defaultwidth = "four";
	this.minimumheight = "five";
	this.minimumwidth = "six";
	this.pagestatus = "cancel";
}

function CAttributeSearchObj()
{
	this.AID = "";
	this.AIDName = "";
	this.AIDCaption = "";
	this.AIDInputType = "";	
	this.Selected = false;
}
function COverRideObj()
{
	this.attr_instance_id = "";
	this.name = "";
}

var OptionsObj = new CMyOptions();
var SelectedObj = new CItemClass();
var AttributeSearchObj = new CAttributeSearchObj();
var OverRideObj = new COverRideObj();


function BtnProperties_onclick()
{

	var ObjX = parent.frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem();
	if (ObjX != null)
	 {
	SelectedObj.SampleValue = ObjX.SampleValue;
	SelectedObj.name = ObjX.attributekey;
	SelectedObj.attributekey = ObjX.attributekey;
	SelectedObj.formatstring = ObjX.formatstring;
	SelectedObj.xpos = ObjX.xpos;
	SelectedObj.ypos = ObjX.ypos;
	SelectedObj.width = ObjX.width;
	SelectedObj.height = ObjX.height;
	SelectedObj.label = ObjX.label;
	SelectedObj.sequence = ObjX.sequence;
	SelectedObj.lucolumn_name = ObjX.GetExtraProperty("LUCOLUMN_NAME");
	SelectedObj.ludisplay_flg = ObjX.GetExtraProperty("LUDISPLAY_FLG");
	SelectedObj.lustorage_flg = ObjX.GetExtraProperty("LUSTORAGE_FLG");
	SelectedObj.lustorage_name = ObjX.GetExtraProperty("LUSTORAGE_NAME");
	SelectedObj.type = ObjX.GetExtraProperty("TYPE");
	SelectedObj.mandatory = ObjX.MandatoryFlag;
	
	SelectedObj.attribute_id = ObjX.GetExtraProperty("ATTRIBUTE_ID");
	SelectedObj.caption = ObjX.GetExtraProperty("CAPTION");
	SelectedObj.inputtype = ObjX.GetExtraProperty("INPUTTYPE");
	SelectedObj.entrymask = ObjX.GetExtraProperty("ENTRYMASK");
	SelectedObj.validvaluefield_flg = ObjX.GetExtraProperty("VALIDVALUEFIELD_FLG");
	SelectedObj.defaultvalue = ObjX.GetExtraProperty("DEFAULTVALUE");
	SelectedObj.unknownvalue = ObjX.GetExtraProperty("UNKNOWNVALUE");
	SelectedObj.textlength = ObjX.GetExtraProperty("TEXTLENGTH");
	SelectedObj.visiblerule_id = ObjX.GetExtraProperty("VISIBLERULE_ID");
	SelectedObj.enabledrule_id = ObjX.GetExtraProperty("ENABLEDRULE_ID");
	SelectedObj.validrule_id = ObjX.GetExtraProperty("VALIDRULE_ID");
	SelectedObj.persistrule_id = ObjX.GetExtraProperty("PERSISTRULE_ID");
	SelectedObj.action_id = ObjX.GetExtraProperty("ACTION_ID");
	SelectedObj.spellcheck_flg = ObjX.GetExtraProperty("SPELLCHECK_FLG");
	SelectedObj.reapplyoverride_flg = ObjX.GetExtraProperty("REAPPLYOVERRIDE_FLG");
	SelectedObj.helpstring = ObjX.GetExtraProperty("HELPSTRING");
	SelectedObj.description = ObjX.GetExtraProperty("DESCRIPTION");
	SelectedObj.lu_type_id = ObjX.GetExtraProperty("LU_TYPE_ID");
	SelectedObj.attributeframe_id = ObjX.GetExtraProperty("ATTRIBUTEFRAME_ID");

	lret = window.showModalDialog("CFPropertiesModal.asp?ATTR_INSTANCE_ID=X&?ATTRIBUTE_ID=X", SelectedObj, "dialogWidth=520px; dialogHeight=580px; center=yes");
		
	if (SelectedObj.pagestatus != "cancel")
	 {
	ObjX.UpdateGroupBegin();
	ObjX.attributekey = SelectedObj.attributekey;
	ObjX.SampleValue = SelectedObj.SampleValue;
	ObjX.attributekey = SelectedObj.name;
	ObjX.xpos = SelectedObj.xpos;
	ObjX.ypos = SelectedObj.ypos;
	ObjX.width = SelectedObj.width;
	ObjX.height = SelectedObj.height;
	ObjX.sequence = SelectedObj.sequence;
	ObjX.SetExtraProperty ("LUCOLUMN_NAME", true, SelectedObj.lucolumn_name);
	ObjX.SetExtraProperty ("LUDISPLAY_FLG", true, SelectedObj.ludisplay_flg);
	ObjX.SetExtraProperty ("LUSTORAGE_FLG", true, SelectedObj.lustorage_flg);
	ObjX.SetExtraProperty ("LUSTORAGE_NAME", true, SelectedObj.lustorage_name);
	ObjX.SetExtraProperty ("TYPE", true, SelectedObj.type);
	ObjX.MandatoryFlag = SelectedObj.mandatory;
		
	if (SelectedObj.caption != "-999999999")
	{
	ObjX.label = SelectedObj.caption;
	}
	else
	{
	ObjX.label = ObjX.GetExtraProperty("ATTR_CAP");
	}
	
	ObjX.SetExtraProperty ("CAPTION", true, SelectedObj.caption);
	ObjX.SetExtraProperty ("INPUTTYPE", true, SelectedObj.inputtype);
	ObjX.SetExtraProperty ("ENTRYMASK", true, SelectedObj.entrymask);
	ObjX.SetExtraProperty ("VALIDVALUEFIELD_FLG", true, SelectedObj.validvaluefield_flg);
	ObjX.SetExtraProperty ("DEFAULTVALUE", true ,SelectedObj.defaultvalue);
	ObjX.SetExtraProperty ("UNKNOWNVALUE", true, SelectedObj.unknownvalue);
	ObjX.SetExtraProperty ("TEXTLENGTH", true, SelectedObj.textlength);
	ObjX.SetExtraProperty ("VISIBLERULE_ID", true, SelectedObj.visiblerule_id);
	ObjX.SetExtraProperty ("ENABLEDRULE_ID", true, SelectedObj.enabledrule_id);
	ObjX.SetExtraProperty ("VALIDRULE_ID", true, SelectedObj.validrule_id);
	ObjX.SetExtraProperty ("PERSISTRULE_ID", true, SelectedObj.persistrule_id);
	ObjX.SetExtraProperty ("ACTION_ID", true, SelectedObj.action_id);
	ObjX.SetExtraProperty ("SPELLCHECK_FLG", true, SelectedObj.spellcheck_flg);
	ObjX.SetExtraProperty ("REAPPLYOVERRIDE_FLG", true, SelectedObj.reapplyoverride_flg);
	ObjX.SetExtraProperty ("HELPSTRING", true, SelectedObj.helpstring);
	ObjX.SetExtraProperty ("DESCRIPTION", true, SelectedObj.description);
	ObjX.SetExtraProperty ("LU_TYPE_ID", true, SelectedObj.lu_type_id);
	ObjX.SetExtraProperty ("ATTRIBUTEFRAME_ID", true, SelectedObj.attributeframe_id);
	ObjX.UpdateGroupEnd();
	parent.frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems();
	 }
  }
}

function BtnOverride_onclick() {
var ObjX = parent.frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem();
if (ObjX != null)
{
	if (ObjX.refid != "" && isNaN(ObjX.refid)==false )
	{
		OverRideObj.attr_instance_id = ObjX.refid;
		OverRideObj.name = ObjX.attributekey
		lret = window.showModalDialog("AttributeOVerrideModal.asp?ATTR_INSTANCE_ID=" + ObjX.refid , OverRideObj, "dialogWidth:700px;dialogHeight:375px");
	}
	else
	{
	alert ("Please save attribute instances before adding an over ride.");
	}
}
}
//-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR=#d6cfbd  topmargin=0 leftmargin=0 bottommargin=0 rightmargin=0 CanDocUnloadNowInf="YES">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Attribute
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<FIELDSET STYLE="BORDER-BOTTOM-WIDTH:1;BORDER-TOP-WIDTH:0;BORDER-LEFT-WIDTH:0;BORDER-COLOR:#006699">
<TABLE CELLPADDING=1  cellspacing=0>
<TR><TD CLASS=LABEL>&nbsp;Name:</TD>
<TD CLASS=LABEL>
<SELECT CLASS=LABEL NAME="ListAttributes" STYLE="WIDTH:330">
</SELECT>
</TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\SearchIcon.gif" id=BtnAddAttribute name=BtnAddAttribute STYLE="CURSOR:HAND" TITLE="Attribute Search"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\PropertiesIcon.gif" TITLE="Properties" STYLE="CURSOR:HAND" ID=BtnProperties Name=BtnProperties LANGUAGE=javascript onclick="return BtnProperties_onclick()"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\saveIcon.gif" TITLE="Save" NAME="BtnSave" ID="BtnSave" STYLE="CURSOR:HAND"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\undoIcon.gif" TITLE="Undo" NAME="UNDO" ID="UNDO" STYLE="CURSOR:HAND"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\ZoomIcon.gif" TITLE="Zoom" NAME="BtnZoom" ID="BtnZoom" STYLE="CURSOR:HAND"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\OverRide.gif" TITLE="Over Ride" NAME=BtnOverRide STYLE="CURSOR:HAND" ID="BtnOverRide" LANGUAGE=javascript onclick="return BtnOverride_onclick()"></TD>
<TD CLASS=SPACING WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\deleteicon.gif" TITLE="Remove" STYLE="CURSOR:HAND" id=BtnDelete name=BtnDelete></TD>
<TD WIDTH=30 ALIGN=CENTER><NOBR><SPAN ID="StatusSpan" CLASS=LABEL></SPAN></TD>
</TR>
</TABLE>
</FIELDSET>
<FORM NAME="SAVEDATA"  TARGET="HIDDENPAGE" ACTION="CFSave.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>" METHOD="POST">
<INPUT TYPE="HIDDEN" NAME="TxtUpdateData">
<INPUT TYPE="HIDDEN" NAME="TxtInsertData">
<INPUT TYPE="HIDDEN" NAME="TxtDeleteData">
<INPUT TYPE="HIDDEN" NAME="UPCOUNT">
<INPUT TYPE="HIDDEN" NAME="INCOUNT">
<INPUT TYPE="HIDDEN" NAME="DELCOUNT">
<INPUT TYPE="HIDDEN" NAME="Refresh">
</FORM>
</BODY>
</HTML>
