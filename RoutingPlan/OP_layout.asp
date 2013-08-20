<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<% 

If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "OP_Layout.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

If Len(Request.QueryString("OPID")) < 1 OR IsNumeric(Request.QueryString("OPID")) = False Then
	Session("ErrorMessage") = "On page " &  Request.ServerVariables("SCRIPT_NAME") & " QueryString OPID was Null or Not Numeric"
	Response.Redirect "..\directerror.asp"
End If

Function ConBool ( Field )
	Select Case Field
		Case "N"
			ConBool = false
		Case "Y"
			ConBool = true
		Case Else
			ConBool = false
	End Select
End Function


Function ReplaceStr(TextIn, SearchStr , Replacement)
    
	Dim WorkText, Pointer
    
    WorkText = TextIn
    Pointer = InStr(1, WorkText, SearchStr)
    Do While Pointer > 0
      WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
      Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr)
    Loop
    ReplaceStr = WorkText
End Function
%> 
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
<!--#include file="..\lib\Help.asp"-->

Dim Zoom_Toggle, Dirty_Flag
Dirty_Flag = 0
Zoom_Toggle = 0
Dim SequenceNumber 
SequenceNumber = 0


Sub BtnAdd_onclick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
AttributeSearchObj.Selected = false
	showModalDialog  "../Attribute/AttributeMaintenance.asp?SELECTONLY=TRUE"  , AttributeSearchObj

If AttributeSearchObj.Selected <> false Then
	listarray = split(AttributeSearchObj.AIDName, "||")
	For i = 0 to Ubound(listarray) step 1
		SequenceNumber = SequenceNumber + 1
		Set objCol = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.PageItems
		Set NewObj = objCol.AddItem (SequenceNumber, listarray(i), "X", Xinc,Yinc, 600, 94,"Arial",8,False,False, False, False, False, 0, "TYPE", "NEW", "", False)
		NewObj.SetExtraProperty "BMP", true, ""
		Xinc= Xinc + 100
		Yinc =Yinc + 100
		Set objOption = document.createElement("option")
		objOption.value = SequenceNumber
		objOption.Text = listarray(i)
		ListAttributes.add( objOption )
	Next

	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.SetSelectedItem(SequenceNumber)
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems
End If
End Sub


Function ConDig ( Field )
	Select Case Field
		Case "False"
			ConDig = "N"
		Case "True"
			ConDig = "Y"
	End Select
End Function

Dim Xinc
Dim Yinc
Xinc = 0
Yinc = 0

Function SearchResults (NodeKey, NodeText)
	SequenceNumber = SequenceNumber + 1
	Set objCol = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.PageItems
	Set NewObj = objCol.AddItem (SequenceNumber, NodeText, "X", Xinc,Yinc, 600, 94,"Arial",8,False,False, False, False, False, 0, "TYPE", "NEW", "", False)
	NewObj.SetExtraProperty "BMP", true, ""
	'objCol.AddItem Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.NextSequenceNumber, NodeText, NodeText, Xinc,Yinc, 600, 94,"Arial",8,False,False, False, False, False, 0, "TYPE", "NEW", "", False
	'refID, attrKey,Label ,xPos, yPos, Width , Height , FontName , fontPointSz, MultiLine , Strikeout , Underline , Italic, Bold , Sequence,  ItemType , FormatString , ReadOnly
	Xinc= Xinc + 100
	Yinc =Yinc + 100
	Set objOption = document.createElement("option")
	objOption.value = SequenceNumber
	objOption.Text = NodeText
	ListAttributes.add( objOption )
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.SetSelectedItem(SequenceNumber)
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems
End Function

Sub BtnSave_onclick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
sResult = ""
InsResult = ""
delResult = ""
UpCount = 0
InCount = 0
delCount = 0
	Set objCol = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.PageItems
	For Each x In objCol
		If x.dirty = "True" AND x.Status = "MODIFY" AND x.deleted = false Then
			UpCount = UpCount + 1
			sResult = sResult & "UPDATE OUTPUT_FIELD SET "
			sResult = sResult & "NAME='" &  x.AttributeKey & "',"
			sResult = sResult & "XPos=" & x.XPos & ","
			sResult = sResult & "YPos=" & x.YPos & ","
			sResult = sResult & "Width=" & x.Width & ","
			sResult = sResult & "Height=" & x.Height & ","
			sResult = sResult & "BOLD_FLG='" & Condig(x.bold) & "',"
			sResult = sResult & "ITALIC_FLG='" & ConDig(x.italic) & "',"
			sResult = sResult & "UNDERLINE_FLG='" & ConDig(x.underline) & "',"
			sResult = sResult & "STRIKEOUT_FLG='" & ConDig(x.strikeout) & "',"
			sResult = sResult & "MULTILINE_FLG='" & ConDig(x.multiline) & "',"
			sResult = sResult & "MAPPING='" & Replace(x.formatstring, "'", "''") & "',"
			sResult = sResult & "FONT_NAME='" 	& x.FontName & "',"
			sResult = sResult & "BMP='" 	& x.GetExtraProperty("BMP") & "',"
			sResult = sResult & "FONT_SIZE=" & x.FontPointSize
			sResult = sResult & " WHERE OUTPUT_FLD_ID=" & x.RefID
			sResult = sResult & "|" & VbCrLf
		End If
		
		If X.Dirty AND x.Status = "NEW"  Then
			InCount = InCount + 1
			InsResult = InsResult & "INSERT INTO OUTPUT_FIELD "
			InsResult = InsResult & "( NAME, MAPPING, BOLD_FLG, ITALIC_FLG, MULTILINE_FLG, UNDERLINE_FLG, STRIKEOUT_FLG, XPOS, YPOS, WIDTH, HEIGHT,BMP"
			InsResult = InsResult & ",FONT_NAME, FONT_SIZE, OUTPUT_PAGE_ID, OUTPUT_FLD_ID ) "
			InsResult = InsResult & "VALUES ( '" & UCASE(x.attributeKey) & "',"
			InsResult = InsResult & "'" & Replace(x.formatstring, "'", "''") & "',"
			InsResult = InsResult & "'" & ConDig(x.bold) & "',"
			InsResult = InsResult & "'" & ConDig(x.italic) & "',"
			InsResult = InsResult & "'" & ConDig(x.multiline) & "',"
			InsResult = InsResult & "'" & ConDig(x.underline) & "',"
			InsResult = InsResult & "'" & ConDig(x.strikeout) & "',"
			InsResult = InsResult & x.XPos & ","
			InsResult = InsResult & x.YPos & ","
			InsResult = InsResult & x.Width & ","
			InsResult = InsResult & x.Height & ","
			InsResult = InsResult & "'" & x.GetExtraProperty("BMP") & "',"
			InsResult = InsResult & "'" & x.FontName & "',"
			InsResult = InsResult & x.FontPointSize  & ","
			InsResult = InsResult & "<%= Request.QueryString("OPID") %>"  & ","
			InsResult = InsResult & "|" & VbCrLf
			'InsResult = InsResult & x.RefID & ")"
			x.status = "MODIFY"
		End if
			
		If x.deleted = True and x.status = "DELETED" Then
			x.status = "DELETED"
			delCount = delCount + 1
			delResult = delResult & "delete from output_field where output_fld_id = " & x.refid & " |" & VbCrLf
		End if	
	Next
	If UpCount > 0 OR InCount > 0 OR DelCount > 0 Then
		document.all.TxtUpdateData.Value = sResult
		document.all.TxtInsertData.Value = InsResult
		document.all.TxtDeleteData.Value = delResult
		document.all.UpCOUNT.Value = UpCount
		document.all.InCOUNT.Value = InCount
		document.all.DELCOUNT.Value = DelCount
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.CleanAllDirty
		document.all.SAVEDATA.Submit()
		StatusSpan.innerHTML = "Saved Successfully"
		StatusSpan.style.color = "MAROON"
		If InCount > 0 Then
			'location.href = "OP_LAYOUT.asp?OPID=<%= Request.QueryString("OPID") %>"
			Parent.frames("LAYOUTAREA").location.href = "LayoutBottom.asp?OPID=<%= Request.QueryString("OPID") %>"
		End If
	Else
		StatusSpan.innerHTML = "Nothing to Save"
		StatusSpan.style.color = "MAROON"
	End If
End Sub

Sub BtnZoom_onclick
	If Zoom_Toggle = 0 Then
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.ZoomIn = True
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.WIDTH=1600
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.HEIGHT=2080
		ZoomMenuText = "Zoom Out"
		Zoom_Toggle = 1
	Else
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.ZoomIn = False
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.WIDTH=800
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.HEIGHT=1040
		ZoomMenuText = "Zoom In"
		Zoom_Toggle = 0
	End If
	Parent.Frames("LAYOUTAREA").SetMenu(ZoomMenuText)
End Sub

Sub BtnDelete_onclick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
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

Sub BOLDBUTTON_onclick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
Set SelX = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem()
If IsObject(SelX) Then
	If BOLDBUTTON.className = "UPMENU" Then
		SelX.bold = "true"
		BOLDBUTTON.className = "CLICKMENU"
	Else
		SelX.bold = "false"
		BOLDBUTTON.className = "UPMENU"
	End If
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.redrawitems
End If
End Sub

Sub ITALICBUTTON_onclick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
Set SelX = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem()
If IsObject(SelX) Then
	If ITALICBUTTON.className = "UPMENU" Then
		SelX.italic = "true"
		ITALICBUTTON.className = "CLICKMENU"
	Else
		SelX.italic = "false"
		ITALICBUTTON.className = "UPMENU"
	End If
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.redrawitems
End If
End Sub

Sub UNDERLINEBUTTON_OnClick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
Set SelX = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem()
If IsObject(SelX) Then
	If UNDERLINEBUTTON.className = "UPMENU" Then
		SelX.underline = "true"
		UNDERLINEBUTTON.className = "CLICKMENU"
	Else
		SelX.underline = "false"
		UNDERLINEBUTTON.className = "UPMENU"
	End If
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.redrawitems
End If
End Sub

Sub STRIKETHROUGHBUTTON_OnCLick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
Set SelX = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem()
If IsObject(SelX) Then
	If STRIKETHROUGHBUTTON.className = "UPMENU" Then
		SelX.strikeout = "true"
		STRIKETHROUGHBUTTON.className = "CLICKMENU"
	Else
		SelX.strikeout = "false"
		STRIKETHROUGHBUTTON.className = "UPMENU"
	End If
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.redrawitems
End If
End Sub

Sub TxtFontSize_onchange
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
	If IsNumeric(TxtFontSize.value) Then
		Set X = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem
		If IsObject(x) Then
			x.FontPointSize = Cint(TxtFontSize.value)
			Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems
		End If
	End If
End Sub

Sub CURFONT_onchange
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
Set X = Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem
	If IsObject(x) Then
		x.Fontname = CurFont.value
		Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems
	End If
End Sub

Sub UNDO_onclick
	Parent.Frames("LAYOUTAREA").document.all.LayoutCtl.Undo
End Sub

Sub ListAttributes_OnChange
	Parent.Frames("LAYOUTAREA").LayoutCtl.SetSelectedItem(ListAttributes.Value)
	Set X = Parent.Frames("LAYOUTAREA").LayoutCtl.GetSelectedItem
	If IsObject(X) Then
		XPIX = Parent.Frames("LAYOUTAREA").LayoutCtl.GetXPixelPos(x.xpos) 
		YPIX = Parent.Frames("LAYOUTAREA").LayoutCtl.GetYPixelPos(x.ypos) 
		Parent.Frames("LAYOUTAREA").window.scrollto XPIX, YPIX
	End If
End Sub


Sub document_onkeydown
	if window.event.ctrlKey then
		Select Case Chr(window.event.keyCode)
			Case "Q"
				Parent.Frames("LAYOUTAREA").location.href = "LAYOUTBOTTOM.ASP?OPID=<%= Request.QueryString("OPID") %>&LANDSCAPE=<%=Request.querystring("LANDSCAPE")%>"
		End Select 
	End If
End Sub

Sub BtnNewAtt_onclick
<% If MODE="RO" Then Response.write(" Exit Sub " ) %>
	lret = SearchResults (SequenceNumber, "New")
End Sub

-->
</SCRIPT>

<SCRIPT LANGUAGE=javascript>
<!--
function CItem()
{
	this.SampleValue = "One";
	this.name = "Two";
	this.attributekey = "Three";
	this.formatstring = "Four";
	this.bold = "Five";
	this.italic = "Six";
	this.multiline = "Seven";
	this.underline = "Eight";
	this.strikeout = "nine";
	this.xpos = "ten";
	this.ypos = "eleven";
	this.width = "twelve";
	this.height = "thirteen";
	this.fontname = "fourteen";
	this.fontpointsize = "fifteen";
	this.itemtype = "sixteen";
	this.bmp = "";
	this.pagestatus = "seventeen";
			
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

var OptionsObj = new CMyOptions();
var SelectedObj = new CItem();
var AttributeSearchObj = new CAttributeSearchObj();

function BtnProperties_onclick()
{
	var ObjX = parent.frames("LAYOUTAREA").document.all.LayoutCtl.GetSelectedItem();
	if (ObjX != null)
	 {
	SelectedObj.SampleValue = ObjX.SampleValue;
	SelectedObj.name = ObjX.attributekey;
	SelectedObj.attributekey = ObjX.attributekey;
	SelectedObj.formatstring = ObjX.formatstring;
	SelectedObj.bold = ObjX.bold;
	SelectedObj.italic = ObjX.italic;
	SelectedObj.multiline = ObjX.multiline;
	SelectedObj.underline = ObjX.underline;
	SelectedObj.strikeout = ObjX.strikeout;
	SelectedObj.xpos = ObjX.xpos;
	SelectedObj.ypos = ObjX.ypos;
	SelectedObj.width = ObjX.width;
	SelectedObj.height = ObjX.height;
	SelectedObj.fontname = ObjX.fontname;
	SelectedObj.fontpointsize = ObjX.fontpointsize;
	SelectedObj.itemtype = ObjX.itemtype;
	SelectedObj.bmp = ObjX.GetExtraProperty("BMP");
	
	lret = window.showModalDialog("ODPropertiesModal.asp", SelectedObj, "dialogWidth=450px; dialogHeight=450px; center=yes");
		
	if (SelectedObj.pagestatus != "cancel")
	 {
	ObjX.UpdateGroupBegin();
	ObjX.SampleValue = SelectedObj.SampleValue;
	ObjX.attributekey = SelectedObj.name;
	ObjX.formatstring = SelectedObj.formatstring;
	ObjX.bold = SelectedObj.bold;
	ObjX.italic = SelectedObj.italic;
	ObjX.multiline = SelectedObj.multiline;
	ObjX.underline = SelectedObj.underline;
	ObjX.strikeout = SelectedObj.strikeout;
	ObjX.xpos = SelectedObj.xpos;
	ObjX.ypos = SelectedObj.ypos;
	ObjX.width = SelectedObj.width;
	ObjX.height = SelectedObj.height;
	ObjX.fontname = SelectedObj.fontname;
	ObjX.fontpointsize = SelectedObj.fontpointsize;
	ObjX.itemtype = SelectedObj.itemtype;
	ObjX.SetExtraProperty ("BMP", true, SelectedObj.bmp);
	ObjX.UpdateGroupEnd();
	parent.frames("LAYOUTAREA").document.all.LayoutCtl.RedrawItems();
	 }
  }
}

function OptionsMenu()
{

OptionsObj.defaultfont = parent.frames("LAYOUTAREA").document.all.LayoutCtl.DefaultFontName;
OptionsObj.defaultfontsize = parent.frames("LAYOUTAREA").document.all.LayoutCtl.DefaultFontPointSize;
OptionsObj.defaultwidth = parent.frames("LAYOUTAREA").document.all.LayoutCtl.DefaultWidth;
OptionsObj.defaultheight = parent.frames("LAYOUTAREA").document.all.LayoutCtl.DefaultHeight;
OptionsObj.minimumheight = parent.frames("LAYOUTAREA").document.all.LayoutCtl.MinHeight;
OptionsObj.minimumwidth = parent.frames("LAYOUTAREA").document.all.LayoutCtl.MinWidth;
	
	
	lret = window.showModalDialog("OptionsModal.asp", OptionsObj, "dialogWidth=320px; dialogHeight=200px; center=yes");

if (OptionsObj.pagestatus != "cancel")
{
parent.frames("LAYOUTAREA").document.all.LayoutCtl.DefaultFontName = OptionsObj.defaultfont;
parent.frames("LAYOUTAREA").document.all.LayoutCtl.DefaultFontPointSize = OptionsObj.defaultfontsize;
parent.frames("LAYOUTAREA").document.all.LayoutCtl.defaultwidth = OptionsObj.defaultwidth;
parent.frames("LAYOUTAREA").document.all.LayoutCtl.defaultheight = OptionsObj.defaultheight;
parent.frames("LAYOUTAREA").document.all.LayoutCtl.MinWidth	 = OptionsObj.minimumwidth;
parent.frames("LAYOUTAREA").document.all.LayoutCtl.MinHeight = OptionsObj.minimumheight;
}
}

//-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR="#d6cfbd"  topmargin="0" rightmargin="0" leftmargin="0" bottommargin="0" CanDocUnloadNowInf="YES">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10>&nbsp&#187 Output Page:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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
<TABLE cellspacing=0>
<TR>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\SearchIcon.gif" id=BtnAdd name=BtnAdd STYLE="CURSOR:HAND" TITLE="Attribute Search"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\NewAttribute.gif" id=BtnNewAtt name=BtnNewAtt STYLE="CURSOR:HAND" TITLE="New Output Field"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\PropertiesIcon.gif" TITLE="Output Field Properties" STYLE="CURSOR:HAND" ID=BtnProperties Name=BtnProperties LANGUAGE=javascript onclick="return BtnProperties_onclick()"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\saveIcon.gif" TITLE="Save" NAME="BtnSave" ID="BtnSave" STYLE="CURSOR:HAND"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\undoIcon.gif" TITLE="Undo" NAME="UNDO" ID="UNDO" STYLE="CURSOR:HAND"></TD>
<TD WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\ZoomIcon.gif" TITLE="Zoom" NAME="BtnZoom" ID="BtnZoom" STYLE="CURSOR:HAND"></TD>
<TD CLASS=SPACING WIDTH=30 ALIGN=CENTER><IMG SRC="..\IMAGES\deleteicon.gif" TITLE="Remove Output Field" STYLE="CURSOR:HAND" id=BtnDelete name=BtnDelete></TD>
<TD CLASS=LABEL ><FONT COLOR=MAROON><SPAN ID="StatusSpan" CLASS=LABEL></SPAN></FONT></TD>
</TR>
</TABLE>
</FIELDSET>

<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10>&nbsp&#187 Selected:</TD>
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
<TD CLASS=LABEL>
<SELECT NAME="CURFONT" STYLE="WIDTH:125" CLASS=LABEL>
<OPTION VALUE="TIMES NEW ROMAN">Times New Roman
<OPTION value="Courier New">Courier New
<OPTION VALUE="ARIAL">Arial
</SELECT></TD>
<TD CLASS=LABEL VALIGN=BOTTOM>
<SELECT CLASS=LABEL id=TxtFontSize name=TxtFontSize>
<OPTION VALUE="6">6
<OPTION VALUE="7">7
<OPTION VALUE="8">8
<OPTION VALUE="9">9
<OPTION VALUE="10">10
<OPTION VALUE="12">12
<OPTION VALUE="14">14
<OPTION VALUE="16">16
<OPTION VALUE="18">18
<OPTION VALUE="20">20
<OPTION VALUE="22">22
<OPTION VALUE="24">24
<OPTION VALUE="26">26
<OPTION VALUE="28">28
<OPTION VALUE="36">36
<OPTION VALUE="48">48
<OPTION VALUE="72">72
</SELECT>
</TD>
<TD CLASS=UPMENU ALIGN=CENTER ID=BOLDBUTTON STYLE="WIDTH:14">B</TD>
<TD CLASS=UPMENU ALIGN=CENTER ID=ITALICBUTTON STYLE="WIDTH:14"><I>I</I></TD>
<TD CLASS=UPMENU ALIGN=CENTER ID=UNDERLINEBUTTON STYLE="WIDTH:14"><U>U</U></TD>
<TD CLASS=UPMENU ALIGN=CENTER ID=STRIKETHROUGHBUTTON STYLE="WIDTH:14">S</TD>
</TR>
</TABLE>
</FIELDSET>
<FORM NAME="SAVEDATA" TARGET="hiddenPage" ACTION="Save.asp?OPID=<%= Request.QueryString("OPID") %>" METHOD="POST">
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
