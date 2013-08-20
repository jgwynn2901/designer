<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<% 
If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then
	Session("NAME") = ""
	Response.Redirect "LayoutBottom.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

Function ReplaceQts( InData )
If IsNull(InData) Then
	ReplaceQts = ""
Else
	ReplaceQts = Replace(InData, """" , """""")
End If 
End Function

Function ConBool ( Field )
	Select Case Field
		Case "N"
			ConBool = False
		Case "Y"
			ConBool = True
		Case Else
			ConBool = False
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
<script src="..\lib\frameSync.js"></script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim Hide_Toggle,Switch_Toggle, G_Zoom_Text, G_ShowSelected_Text, G_ShowLabel_text, G_SequenceNumber
Hide_Toggle = 0
Switch_Toggle = 0
G_Zoom_Text = "Zoom In"
G_ShowSelected_Text = "Show Selected"
G_ShowLabel_text = "Show Sample"
G_SequenceNumber= 400

Sub window_onunload
	LayoutCtl.PageItems.RemoveAll
	LayoutCtl.PageItems.RemoveAllItemTemplates
End Sub


Sub Window_OnLoad
Set objCol = LayoutCtl.PageItems
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = "SELECT OUTPUT_PAGE.OUTPUT_PAGE_ID, OUTPUT_FIELD.NAME, OUTPUT_PAGE.PAGE_NUMBER, OUTPUT_PAGE.OUTPUT_TRAY, OUTPUT_PAGE.BACKGROUND_BMP,OUTPUT_PAGE.ORIENTATION, OUTPUT_FIELD.XPOS, OUTPUT_FIELD.YPOS,"
	SQLST = SQLST & "OUTPUT_FIELD.WIDTH, OUTPUT_FIELD.HEIGHT, OUTPUT_FIELD.BMP, OUTPUT_FIELD.FONT_NAME, OUTPUT_FIELD.OUTPUT_FLD_ID, OUTPUT_FIELD.FONT_SIZE, OUTPUT_FIELD.BOLD_FLG, OUTPUT_FIELD.ITALIC_FLG, 	OUTPUT_FIELD.UNDERLINE_FLG,"
	SQLST = SQLST & "OUTPUT_FIELD.MAPPING, OUTPUT_FIELD.STRIKEOUT_FLG, OUTPUT_FIELD.MULTILINE_FLG FROM OUTPUT_PAGE, OUTPUT_FIELD WHERE OUTPUT_PAGE.OUTPUT_PAGE_ID = " & Request.QueryString("OPID") & " AND OUTPUT_PAGE.OUTPUT_PAGE_ID = OUTPUT_FIELD.OUTPUT_PAGE_ID" 
	Set RS = Conn.Execute(SQLST)
	Do While Not RS.EOF
	MULTILINE_FLG = ConBool(RS("MULTILINE_FLG"))
	UNDERLINE_FLG = ConBool(RS("UNDERLINE_FLG"))
	BOLD_FLG = ConBool(RS("BOLD_FLG"))
	ITALIC_FLG = ConBool(RS("ITALIC_FLG"))
	STRIKEOUT_FLG = ConBool(RS("STRIKEOUT_FLG"))
	VB_MAPPING = ReplaceStr( ReplaceStr (RS("MAPPING"), """" , """"""), VBCrlf, "")
%>
	
	Set NewObj = objCol.AddItem ("<%= RS("OUTPUT_FLD_ID") %>", "<%= RS("NAME") %>",  "X", <%= RS("XPOS") %>, <%= RS("YPOS") %>, <%= RS("WIDTH") %>, <%= RS("HEIGHT") %>, "<%= RS("FONT_NAME") %>", <%= RS("FONT_SIZE") %>, <%= MULTILINE_FLG %>, <%= STRIKEOUT_FLG %>, <%= UNDERLINE_FLG %>, <%= ITALIC_FLG %>,<%=  BOLD_FLG %>, 0, "Output Field", "MODIFY",  "<%= VB_MAPPING %>", False	)
	<% If MODE = "RO" Then %>
	NewObj.readonly = true
	<% End If %>
	
	NewObj.SetExtraProperty "BMP", true, "<%= ReplaceQts(RS("BMP").value) %>"
<%
	'refID, attrKey,Label ,xPos, yPos, Width , Height , FontName , fontPointSz, MultiLine , Strikeout , Underline , Italic, Bold , Sequence,  ItemType , FormatString , ReadOnly
	RS.MoveNext
	Loop
	If Not RS.EOF Then
		RS.MoveFirst
	Else
		SQL1 = ""
		SQL1 = SQL1 & "SELECT BACKGROUND_BMP , ORIENTATION FROM OUTPUT_PAGE WHERE OUTPUT_PAGE_ID = " & Request.QueryString("OPID")
		Set RS2 = Conn.Execute(SQL1)
		If Not RS2.EOF Then
		   BACKGROUND_BMP = RS2("BACKGROUND_BMP")
		   ORIENTATION    = RS2("ORIENTATION")
		   if ORIENTATION = "L" then 
              XLogicalExtent = 6324
              YLogicalExtent = 4800
           else
              XLogicalExtent = 4800
              YLogicalExtent = 6324
           end if

		End If
	End If
%>	

'LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|Zoom|100|COMMAND|Show All|125|COMMAND|Switch Labels|155|COMMAND|Options...|101|COMMAND|Save|110", "|"
LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Adjust All|651|COMMAND|Options...|101|COMMAND|Save|110", "|"
LayoutCtl.SetContextMenu "FIELD", "COMMAND|Properties...|400|COMMAND|Duplicate|402|SEPARATOR|0|0|COMMAND|Remove|401", "|" 
LayoutCtl.CleanAllDirty
LayoutCtl.EnableTrackUndo = True
LayoutCtl.RedrawItems

While (Parent.Frames("LAYOUTAREA").LayoutCtl.readyState <> 4)
	temp = 1
Wend

wait4Frame Parent.Frames("TOPAREA"), "LoadDropDown()"

End Sub

Sub SetMenu(ZoomText)
G_Zoom_Text = ZoomText
LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Adjust All|651|COMMAND|Options...|101|COMMAND|Save|110", "|"
End Sub

Sub RemoveOptions( objSelect )
	Dim intOptIndex
	If objSelect.options.length > 0 Then
		while (objSelect.options.length > 0)
			objSelect.Remove( intOptIndex )
		wend
	End if
End Sub

Function LoadDropDown
If IsObject (Parent.Frames("TOPAREA").document.all.ListAttributes) = True Then
RemoveOptions Parent.Frames("TOPAREA").document.all.ListAttributes
Set objCol = LayoutCtl.PageItems
For Each x In objCol
	Set objOption = Parent.Frames("TOPAREA").document.createElement("option")
	objOption.value = x.refid
	objOption.Text = x.AttributeKey
	Parent.Frames("TOPAREA").document.all.ListAttributes.options.add( objOption )
Next
Else
	MsgBox "Load Drop Down Failed"
End If
End Function

Sub LayoutCtl_OnItemCreated( logXPos , logYPos ,logWidth , logHeight , SequenceNum )
	Set objOption = Parent.Frames("TOPAREA").document.createElement("option")
	Set X=LayoutCtl.GetSelectedItem
	objOption.value = x.RefID
	objOption.Text = x.AttributeKey
	X.SetExtraProperty "BMP", true, ""
	Parent.Frames("TOPAREA").document.all.ListAttributes.options.add( objOption )
End Sub

Sub LayoutCtl_OnMenuCommand (MenuID)
	Select Case Cstr(MenuID)
		Case "100" ' Zoom
			Parent.Frames("TOPAREA").Document.all.BtnZoom.click()
		Case "101" 'Options
			Parent.Frames("TOPAREA").OptionsMenu()
		Case "400" 'Properties
			Parent.Frames("TOPAREA").Document.all.BtnProperties.onclick()
		Case "401" 'Remove
			Parent.Frames("TOPAREA").Document.all.BtnDelete.click()
		Case "110" 'Save
			Parent.Frames("TOPAREA").Document.all.BtnSave.click()
		Case "145" 'undo
			document.all.LayoutCtl.Undo
		Case "155" ' Switch Label
			If Switch_Toggle = 0 Then
				document.all.LayoutCtl.ShowAttribute = false
				document.all.LayoutCtl.ShowSampleValue = true
				Switch_Toggle = 1
				G_ShowLabel_text = "Show Names"
			Else
				document.all.LayoutCtl.ShowSampleValue = false
				document.all.LayoutCtl.ShowAttribute = true
				Switch_Toggle = 0
				G_ShowLabel_text = "Show Sample"
			End if
		LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Adjust All|651|COMMAND|Options...|101|COMMAND|Save|110", "|"
		Case "402" 'Duplicate
			Call DuplicateField
		Case "651" 'Adjust All
			Call AdjustAll()
		Case "125" 'Hide
			If Hide_Toggle = 0 Then
				document.all.LayoutCtl.ShowSelectedOnly = True
				G_ShowSelected_Text = "Show All"
				Hide_Toggle = 1
			Else
				document.all.LayoutCtl.ShowSelectedOnly = False
				G_ShowSelected_Text = "Show Selected"
				Hide_Toggle = 0
			End if
			LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Adjust All|651|COMMAND|Options...|101|COMMAND|Save|110", "|"
	End Select
End Sub

Sub LayoutCtl_OnItemSelect( pixelXPos, pixelYPos, xPos , yPos, attrKey,  refID )
	If attrKey <> "" OR refID <> "" Then
		Set X =LayoutCtl.GetSelectedItem	
		Parent.Frames("TOPAREA").document.all.ListAttributes.value = x.refid
		Parent.Frames("TOPAREA").document.all.CURFONT.value = UCASE(X.FontName)
		Parent.Frames("TOPAREA").document.all.TxtFontSize.value = x.FontPointSize
		If x.underline = "True" Then
			Parent.Frames("TOPAREA").document.all.UNDERLINEBUTTON.className = "CLICKMENU"
		Else
			Parent.Frames("TOPAREA").document.all.UNDERLINEBUTTON.className = "UPMENU"
		End If
		If x.strikeout = "True" Then
			Parent.Frames("TOPAREA").document.all.STRIKETHROUGHBUTTON.className = "CLICKMENU"
		Else
			Parent.Frames("TOPAREA").document.all.STRIKETHROUGHBUTTON.className = "UPMENU"
		End If
		
		If x.bold = "True" Then
			Parent.Frames("TOPAREA").document.all.BOLDBUTTON.className = "CLICKMENU"
		Else
			Parent.Frames("TOPAREA").document.all.BOLDBUTTON.className = "UPMENU"
		End If
		If x.italic = "True" Then
			Parent.Frames("TOPAREA").document.all.ITALICBUTTON.className = "CLICKMENU"
		Else
			Parent.Frames("TOPAREA").document.all.ITALICBUTTON.className = "UPMENU"
		End If
	Else
		Parent.Frames("TOPAREA").document.all.ListAttributes.value = ""
		Parent.Frames("TOPAREA").document.all.CURFONT.value = ""
		Parent.Frames("TOPAREA").document.all.TxtFontSize.value = ""
		Parent.Frames("TOPAREA").document.all.UNDERLINEBUTTON.className = "UPMENU"
		Parent.Frames("TOPAREA").document.all.STRIKETHROUGHBUTTON.className = "UPMENU"
		Parent.Frames("TOPAREA").document.all.BOLDBUTTON.className = "UPMENU"
		Parent.Frames("TOPAREA").document.all.ITALICBUTTON.className = "UPMENU"
	End If
End Sub

Sub DuplicateField
	G_SequenceNumber = G_SequenceNumber + 1
	Set X=LayoutCtl.GetSelectedItem
	Set objCol = LayoutCtl.PageItems
	Set NewObj = objCol.AddItem (G_SequenceNumber, x.attributekey,  "X", x.xpos + x.height, x.ypos + x.height, x.width, x.height, x.fontname, x.fontpointsize, x.Multiline, x.strikeout,  x.underline,  x.italic, x.bold,  0, "TYPE", "NEW",  x.formatstring , False	)
	NewObj.SetExtraProperty "BMP", true, X.GetExtraProperty("BMP")
	Set objOption = Parent.Frames("TOPAREA").document.createElement("option")
	objOption.value = G_SequenceNumber
	objOption.Text = x.attributekey
	Parent.Frames("TOPAREA").document.all.ListAttributes.options.add( objOption )
	LayoutCtl.SetSelectedItem(G_SequenceNumber)
	LayoutCtl.redrawitems
End Sub

function AdjustAll()
AdjustAllObj.Adjust_X = ""
AdjustAllObj.Adjust_Y = ""
AdjustAllObj.Adjust_Height = ""
AdjustAllObj.Adjust_Width = ""
lret = window.showModalDialog("AdjustAllModal.asp", AdjustAllObj, "dialogWidth=250px; dialogHeight=200px; center=yes")
Set objCol = LayoutCtl.PageItems
If AdjustAllObj.Adjust_X <> "" AND IsNumeric(AdjustAllObj.Adjust_X)Then
	For Each x In objCol
	x.xPos = x.xPos + AdjustAllObj.Adjust_X
	Next
End If

If AdjustAllObj.Adjust_Y <> "" AND IsNumeric(AdjustAllObj.Adjust_Y) Then
	For Each x In objCol
	x.yPos = x.yPos + AdjustAllObj.Adjust_Y
	Next
End If

If AdjustAllObj.Adjust_Height <> "" AND IsNumeric(AdjustAllObj.Adjust_Height) Then
	For Each x In objCol
	x.height = x.height + AdjustAllObj.Adjust_Height
	Next
End If

If AdjustAllObj.Adjust_Width <> "" AND IsNumeric(AdjustAllObj.Adjust_Width)Then
	For Each x In objCol
	x.Width = x.Width + AdjustAllObj.Adjust_Width
	Next
End if
LayoutCtl.RedrawItems
End Function
-->
</SCRIPT>
<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
	if (false == confirm("Leave Page without saving?"))
		return false;
	else
		return true;
}
function CAdjustAllObj()
{
	this.Adjust_X = "";
	this.Adjust_Y = "";
	this.Adjust_Height = "";
	this.Adjust_Width = "";
	this.pagestatus = "cancel";
}
var AdjustAllObj = new CAdjustAllObj();

</script>
</HEAD>
<BODY topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 CanDocUnloadNowInf=YES>

<OBJECT VIEWASTEXT ID="LayoutCtl" 
<!--#include file="..\lib\EditorCLSID.inc"-->
WIDTH='800'  HEIGHT='1040' >
<PARAM NAME="ZoomIn" VALUE="False">
<PARAM NAME="XViewExtent" VALUE="770">
<PARAM NAME="YViewExtent" VALUE="1014">
<PARAM NAME="XLogicalExtent" VALUE="<%= XLogicalExtent%>">
<PARAM NAME="YLogicalExtent" VALUE="<%= YLogicalExtent%>">
<PARAM NAME="DefaultHeight" VALUE="30">
<PARAM NAME="DefaultWidth" VALUE="200">
<PARAM NAME="DefaultBold" VALUE="True">
<PARAM NAME="MinWidth" VALUE="50">
<PARAM NAME="MinHeight" VALUE="50">
<PARAM NAME="DefaultYPos" VALUE="1">
<PARAM NAME="DefaultXPos" VALUE="1">
<PARAM NAME="XSelectedSensitivity" VALUE="50">
<PARAM NAME="YSelectedSensitivity" VALUE="50">
<PARAM NAME="DefaultFontPointSize" VALUE="8">
<PARAM NAME="DefaultFontName" VALUE="Arial">
<PARAM NAME="DefaultItemType" VALUE="Output Field">
<PARAM NAME="DefaultAttributeKey" VALUE="New">
<PARAM NAME="DefaultLabel" VALUE="New">
<PARAM NAME="DefaultBorderColor" VALUE="16711680">
<PARAM NAME="DefaultItemTextColor" VALUE="39219">
<PARAM NAME="DefaultSelectedBorderColor" VALUE="255">
<PARAM NAME="DefaultSampleValue" VALUE="X">
<PARAM NAME="DefaultHighlightedBorderColor" VALUE="16119260">
<PARAM NAME="DefaultBorderThickness" VALUE="10">
<PARAM NAME="BackImageURL" VALUE="HTTP://<%= Request.servervariables("SERVER_NAME") %>/images/BGBMP/<%= BACKGROUND_BMP %>">
<PARAM NAME="NextSequenceNumber" VALUE="1">
<PARAM NAME="UseAnisotropicMapMode" VALUE="False">
<PARAM NAME="UseTransparentItemBackground" VALUE="True">
<PARAM NAME="ShowSampleValue" VALUE="True">
<PARAM NAME="ShowAttribute" VALUE="True">
<PARAM NAME="AutoCreateItem" VALUE="True">
</OBJECT>
</BODY>
</HTML>