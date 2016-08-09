<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<% 
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then   
	Session("NAME") = ""
	Response.Redirect "CF_LayoutBottom.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"

nFrameOffsetY = 5
nExtremeY = 75
nClientY = 400
LayoutCtloffset = Round((nFrameOffsetY/nExtremeY) * nClientY )

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
<META name=VI60_defaultClientScript content=VBScript>
<script src="..\lib\frameSync.js"></script>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
	if (false == confirm("Leave Page without saving?"))
		return false;
	else
		return true;
}

function CAttributeSearchObj()
{
	this.AID = "";
	this.AIDName = "";
	this.AIDCaption = "";
	this.AIDInputType = "";	
}

var AttributeSearchObj = new CAttributeSearchObj();
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim Hide_Toggle,Switch_Toggle, G_Zoom_Text
Dim G_ShowSelected_Text, G_ShowLabel_text, G_SequenceNumber
Hide_Toggle = 0
Switch_Toggle = 0
G_Zoom_Text = "Zoom In"
G_ShowSelected_Text = "Show Selected"
G_ShowLabel_text = "Show Samples"
G_SequenceNumber= 400


Sub window_onunload
	LayoutCtl.PageItems.RemoveAll
	LayoutCtl.PageItems.RemoveAllItemTemplates
End Sub

Sub window_onload

Set objCol = LayoutCtl.PageItems

<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = "SELECT ATTR_INSTANCE.*, ATTRIBUTE.CAPTION As ATTRIBUTE_CAPTION, ATTRIBUTE.ATTRIBUTE_ID, ATTRIBUTE.NAME FROM ATTR_INSTANCE, ATTRIBUTE WHERE ATTR_INSTANCE.ATTRIBUTE_ID = ATTRIBUTE.ATTRIBUTE_ID AND FRAME_ID=" & Request.QueryString("FRAMEID") & " ORDER BY ATTRIBUTE.NAME"
	SQLST2 = SQLST2 & "SELECT TITLE FROM FRAME WHERE FRAME_ID=" & Request.QueryString("FRAMEID") 
	Set RS = Conn.Execute(SQLST)
	Set RS2 = Conn.Execute(SQLST2) 
%> 


			'Templates
			Set NewObj = objCol.AddItemTemplate("CMDBTN", "",  "", 0, 0, 15, 4, "ARIAL", 9,false ,false ,false ,false ,false , 1, "", "MODIFY",  "", False	)
			NewObj.LabelOffsetY = -3
			NewObj.LabelOffsetX = 0
			NewObj.ItemBackColor = 26316
						
			Set NewObj = objCol.AddItemTemplate("!TRIGGER", "",  "", 0, 0, 15, 4, "ARIAL", 9,false ,false ,false ,false ,false , 1, "", "MODIFY",  "", False	)
			NewObj.LabelOffsetY = -3
			NewObj.LabelOffsetX = 0
			NewObj.ItemBackColor = 26316
			
			Set NewObj = objCol.AddItemTemplate("!SUMMARY", "",  "", 0, 0, 15, 4, "ARIAL", 9,false ,false ,false ,false ,false , 1, "", "MODIFY",  "", False	)
			NewObj.LabelOffsetY = -3
			NewObj.LabelOffsetX = 0
			NewObj.ItemBackColor = 26316
			
<% Do While Not RS.EOF %>
<%
If RS("CAPTION") = "-999999999" Then
	UseCaption = ReplaceStr(RS("ATTRIBUTE_CAPTION"), """", """""")
Else
	UseCaption = ReplaceStr(RS("CAPTION"), """", """""")
End If
%>

			Set NewObj = objCol.AddItem ("<%= RS("ATTR_INSTANCE_ID") %>", "<%= ReplaceStr(RS("NAME"), """", """""")  %>",  "<%= UseCaption %>", <%= RS("XPOS") %>, <%= RS("YPOS") %>, <%= RS("WIDTH") %>, <%= RS("HEIGHT") %>, "ARIAL", 9,false ,false ,false ,false ,false , <%= Clng(RS("SEQUENCE")) %>, "<%= RS("INPUTTYPE") %>", "MODIFY",  "", False	)
			<% If MODE = "RO" Then %>
			NewObj.readonly = true
			<% End If %>
			NewObj.MandatoryFlag = <%= ConBool(RS("MANDATORY_FLG")) %>
			NewObj.SampleValue = "X"
			NewObj.SetExtraProperty "LUCOLUMN_NAME", true, "<%= ReplaceStr(RS("LUCOLUMN_NAME"), """", """""") %>"
			NewObj.SetExtraProperty "LUDISPLAY_FLG", true, <%=ConBool( RS("LUDISPLAY_FLG")) %>
			NewObj.SetExtraProperty "LUSTORAGE_FLG", true, <%= ConBool(RS("LUSTORAGE_FLG")) %>
			NewObj.SetExtraProperty "LUSTORAGE_NAME", true, "<%= ReplaceStr(RS("LUSTORAGE_NAME"), """", """""") %>"
			NewObj.SetExtraProperty "ATTRIBUTE_ID", true, "<%= RS("ATTRIBUTE_ID") %>"
			NewObj.SetExtraProperty "TYPE" , true, "<%= RS("TYPE") %>"
			NewObj.SetExtraProperty "ATTR_CAP", true, "<%= ReplaceStr(RS("ATTRIBUTE_CAPTION"), """", """""") %>"
			NewObj.SetExtraProperty "CAPTION", true, "<%= ReplaceStr(RS("CAPTION"), """", """""") %>"
			NewObj.SetExtraProperty "INPUTTYPE", true, "<%= ReplaceStr(RS("INPUTTYPE"), """", """""") %>"
			NewObj.SetExtraProperty "ENTRYMASK", true, "<%= ReplaceStr(RS("ENTRYMASK"), """", """""") %>"
			NewObj.SetExtraProperty "VALIDVALUEFIELD_FLG", true, "<%= ReplaceStr(RS("VALIDVALUEFIELD_FLG"), """", """""") %>"
			NewObj.SetExtraProperty "DEFAULTVALUE", true, "<%= ReplaceStr(RS("DEFAULTVALUE"), """", """""") %>"
			NewObj.SetExtraProperty "UNKNOWNVALUE", true, "<%= ReplaceStr(RS("UNKNOWNVALUE"), """", """""") %>"
			NewObj.SetExtraProperty "TEXTLENGTH", true, "<%= ReplaceStr(RS("TEXTLENGTH"), """", """""") %>"
			NewObj.SetExtraProperty "VISIBLERULE_ID", true, "<%= ReplaceStr(RS("VISIBLERULE_ID"), """", """""") %>"
			NewObj.SetExtraProperty "ENABLEDRULE_ID", true, "<%= ReplaceStr(RS("ENABLEDRULE_ID"), """", """""") %>"
			NewObj.SetExtraProperty "VALIDRULE_ID", true, "<%= ReplaceStr(RS("VALIDRULE_ID"), """", """""") %>"
			NewObj.SetExtraProperty "PERSISTRULE_ID", true, "<%= ReplaceStr(RS("PERSISTRULE_ID"), """", """""") %>"
			NewObj.SetExtraProperty "ACTION_ID", true, "<%= ReplaceStr(RS("ACTION_ID"), """", """""") %>"
			NewObj.SetExtraProperty "SPELLCHECK_FLG", true, "<%= ReplaceStr(RS("SPELLCHECK_FLG"), """", """""") %>"
			NewObj.SetExtraProperty "REAPPLYOVERRIDE_FLG", true, "<%= ReplaceStr(RS("REAPPLYOVERRIDE_FLG"), """", """""") %>"
			NewObj.SetExtraProperty "HELPSTRING", true, "<%= ReplaceStr(RS("HELPSTRING"), """", """""") %>"
			NewObj.SetExtraProperty "DESCRIPTION", true, "<%= ReplaceStr(RS("DESCRIPTION"), """", """""") %>"
			NewObj.SetExtraProperty "LU_TYPE_ID", true, "<%= ReplaceStr(RS("LU_TYPE_ID"), """", """""") %>"
			<% If IsNull(RS("ATTRIBUTEFRAME_ID")) Then
				ATTRIBUTEFRAME_ID = "null"
			Else
				ATTRIBUTEFRAME_ID = RS("ATTRIBUTEFRAME_ID") 
			End If %>
			
			NewObj.SetExtraProperty "ATTRIBUTEFRAME_ID", true, "<%= ReplaceStr(ATTRIBUTEFRAME_ID, """", """""") %>"
<%
RS.MoveNext
Loop	
%>
LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Save|110", "|"
LayoutCtl.SetContextMenu "FIELD", "COMMAND|Instance Properties...|400|COMMAND|Attribute Properties...|500|COMMAND|Over Ride|423|SEPARATOR|0|0|COMMAND|Remove|401", "|" 
LayoutCtl.CleanAllDirty
LayoutCtl.EnableTrackUndo = True
LayoutCtl.RedrawItems

While (Parent.Frames("LAYOUTAREA").LayoutCtl.readyState <> 4)
	temp = 1
Wend

wait4Frame Parent.Frames("TOPAREA"), "LoadDropDown()"

End Sub

Sub LayoutCtl_OnMenuCommand (MenuID)
	Select Case Cstr(MenuID)
		Case "100" 'Zoom
			Parent.Frames("TOPAREA").Document.all.BtnZoom.click()
		Case "400" 'Instance Properties
			Parent.Frames("TOPAREA").Document.all.BtnProperties.onclick()
		Case "500" 'Attribute Properties
			
			Set X = Parent.Frames("LAYOUTAREA").LayoutCtl.GetSelectedItem
			AttributeSearchObj.AIDCaption = x.label
			AttributeSearchObj.AIDName = x.attributekey
			AttributeSearchObj.AIDInputType = x.ItemType
			
			lret = window.showModalDialog ("../Attribute/AttributeMaintenance.asp?DETAILONLY=TRUE&AID=" & X.GetExtraProperty("ATTRIBUTE_ID") , AttributeSearchObj, "center=yes")
			
			x.label = AttributeSearchObj.AIDCaption
			x.attributekey = AttributeSearchObj.AIDName
			x.ItemType = AttributeSearchObj.AIDInputType
			Parent.Frames("LAYOUTAREA").LayoutCtl.RedrawItems()
		Case "401" 'Remove
		
			Parent.Frames("TOPAREA").Document.all.BtnDelete.click()
		Case "110" 'Save
			Parent.Frames("TOPAREA").Document.all.BtnSave.click()
		Case "145" 'undo
			document.all.LayoutCtl.Undo
		Case "155" 'Switch Label
			If Switch_Toggle = 0 Then
				document.all.LayoutCtl.ShowAttribute = false
				document.all.LayoutCtl.ShowSampleValue = true
				Switch_Toggle = 1
				G_ShowLabel_text = "Show Attributes"
			Else
				document.all.LayoutCtl.ShowAttribute = true
				document.all.LayoutCtl.ShowSampleValue = false
				Switch_Toggle = 0
				G_ShowLabel_text = "Show Samples"
			End if
			LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Options...|101|COMMAND|Save|110", "|"
		Case "402" 'Duplicate
			Call DuplicateField
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
			LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Options...|101|COMMAND|Save|110", "|"
		Case "423" ' Override
			Parent.Frames("TOPAREA").Document.all.BtnOverride.click()
			LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Options...|101|COMMAND|Save|110", "|"
	End Select
End Sub

Sub SetMenu(ZoomText)
	G_Zoom_Text = ZoomText
	LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Options...|101|COMMAND|Save|110", "|"
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
	Set objOption = Parent.Frames("TOPAREA").document.createElement("OPTION")
	objOption.value = x.refid
	objOption.Text = x.AttributeKey
	Parent.Frames("TOPAREA").document.all.ListAttributes.Options.add( objOption )
Next
Else
	MsgBox "Load Drop Down Failed"
End If
	LoadDropDown = True
End Function

Sub LayoutCtl_OnItemSelect( pixelXPos, pixelYPos, xPos , yPos, attrKey,  refID )
	Set X =LayoutCtl.GetSelectedItem	
	If IsObject(X) Then
		Parent.Frames("TOPAREA").document.all.ListAttributes.value = x.refid
	End If
End Sub
-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR=SILVER topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 CanDocUnloadNowInf=YES>
<DIV ID=HEADER style="position:absolute;width:102%;background-color:#006666;color:white;font:bold 12pt">&nbsp;<%= UCASE(RS2("TITLE")) %></DIV><BR>
<OBJECT ID="LayoutCtl" style="POSITION:ABSOLUTE;TOP:<%= LayoutCtloffset %>;LEFT:0;"
<!--#include file="..\lib\EditorCLSID.inc"-->
width=883 height=533>
	<PARAM NAME="DefaultLabel" VALUE="Def. Label">
	<PARAM NAME="DefaultSampleValue" VALUE="X">
	<PARAM NAME="DefaultAttributeKey" VALUE="Caller:Name:LastName">
	<PARAM NAME="DefaultItemType" VALUE="Output Field">
	<PARAM NAME="DefaultFontPointSize" VALUE="10">
	<PARAM NAME="DefaultXPos" VALUE="1">
	<PARAM NAME="DefaultYPos" VALUE="1">
	<PARAM NAME="DefaultWidth" VALUE="5">
	<PARAM NAME="DefaultHeight" VALUE="3">
	<PARAM NAME="DefaultBorderColor" VALUE="16777215">
	<PARAM NAME="DefaultSelectedBorderColor" VALUE="65535">
	<PARAM NAME="XLogicalExtent" VALUE="100">
	<PARAM NAME="YLogicalExtent" VALUE="100">
	<PARAM NAME="ZoomIn" VALUE="0">
	<PARAM NAME="DefaultFormatString" VALUE="Testing FormatString">
	<PARAM NAME="DefaultMultiLine" VALUE="0">
	<PARAM NAME="DefaultStrikeout" VALUE="0">
	<PARAM NAME="DefaultUnderline" VALUE="0">
	<PARAM NAME="DefaultItalic" VALUE="0">
	<PARAM NAME="DefaultBold" VALUE="0">
	<PARAM NAME="MinWidth" VALUE="1">
	<PARAM NAME="MinHeight" VALUE="1">
	<PARAM NAME="XSelectedSensitivity" VALUE="1">
	<PARAM NAME="YSelectedSensitivity" VALUE="1">
	<PARAM NAME="ShowLabel" VALUE="TRUE">
	<PARAM NAME="DefaultMandatoryFlag" VALUE="FALSE">
	<PARAM NAME="DefaultSampleValue" VALUE="X">
	<PARAM NAME="DefaultLabelOffsetX" VALUE="0">
	<PARAM NAME="DefaultLabelOffsetY" VALUE="-3">
	<PARAM NAME="ShowAttribute" VALUE="TRUE">
	<PARAM NAME="ShowSampleValue" VALUE="FALSE">
	<PARAM NAME="ShowLabel" VALUE="TRUE">
	<PARAM NAME="DefaultBorderThickness" VALUE="0">
	<PARAM NAME="DefaultItemBackColor" VALUE="16777215">
	<PARAM NAME="BoldLabelOnMandatory" VALUE="TRUE">
	<PARAM NAME="DefaultItemBackColorSelected" VALUE="65535">
	<PARAM NAME="DefaultItemTextColor" VALUE="0">
	<PARAM NAME="UseTransparentItemBackground" VALUE="FALSE">
	<PARAM NAME="UseAnisotropicMapMode" VALUE="False">
	<PARAM NAME="ShowGridBorder" VALUE="TRUE">
	<PARAM NAME="DefaultMandatoryBackColor" VALUE="255">
	<PARAM NAME="GridBorderX" VALUE="0">
	<PARAM NAME="GridBorderY" VALUE="0">
	<PARAM NAME="GridBorderWidth" VALUE=<%=Session("LayoutCtlWidth")%>>
	<PARAM NAME="GridBorderHeight" VALUE=<%=Session("LayoutCtlHeight")%>>
	<PARAM NAME="DefaultFontName" VALUE="MS Sans Serif">
	<PARAM NAME="XViewExtent" VALUE="883">
	<PARAM NAME="YViewExtent" VALUE="533">
	<PARAM NAME="AutoCreateItem" VALUE="False">
</OBJECT>
</BODY>
</HTML>
<%
If IsObject(RS) Then
	RS.Close
	Set RS = Nothing
End If

If IsObject(RS2) Then
	RS2.Close
	Set RS2 = Nothing
End If

If IsObject(Conn) Then
	Conn.Close
	Set Conn = Nothing
End If
%>