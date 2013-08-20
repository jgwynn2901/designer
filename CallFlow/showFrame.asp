<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<% 

nFrameOffsetY = 5
nExtremeY = 75
nClientY = 400
LayoutCtloffset = Round((nFrameOffsetY/nExtremeY) * nClientY )

Function ConBool ( cField )
	Select Case cField
		Case "N"
			ConBool = false
		Case "Y"
			ConBool = true
		Case Else
			ConBool = false
	End Select
End Function

Function ReplaceStr(cTextIn, SearchStr , Replacement)
	Dim cWorkText, nPointer
	
    cWorkText = cTextIn
    nPointer = InStr(1, cWorkText, SearchStr)
    Do While nPointer > 0
		cWorkText = Left(cWorkText, nPointer - 1) & Replacement & Mid(cWorkText, nPointer + Len(SearchStr))
		nPointer = InStr(nPointer + Len(Replacement), cWorkText, SearchStr)
    Loop
    ReplaceStr = cWorkText
End Function
%> 
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onunload
	LayoutCtl.PageItems.RemoveAll
	LayoutCtl.PageItems.RemoveAllItemTemplates
End Sub

Sub window_onload
dim oPages

Set oPages = LayoutCtl.PageItems
<%
dim oConn, cSQL, cSQL2, oRS, oRS2, cFrameID

cFrameID = Request.QueryString("FRAMEID")
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
cSQL = "SELECT ATTR_INSTANCE.*, ATTRIBUTE.CAPTION As ATTRIBUTE_CAPTION, ATTRIBUTE.ATTRIBUTE_ID, ATTRIBUTE.NAME FROM ATTR_INSTANCE, ATTRIBUTE WHERE ATTR_INSTANCE.ATTRIBUTE_ID = ATTRIBUTE.ATTRIBUTE_ID AND FRAME_ID=" & Request.QueryString("FRAMEID") & " ORDER BY ATTRIBUTE.NAME"
cSQL2 = "SELECT TITLE FROM FRAME WHERE FRAME_ID=" & Request.QueryString("FRAMEID") 
Set oRS = oConn.Execute(cSQL)

Set oRS2 = oConn.Execute(cSQL2) 
%> 
'Templates
Set NewObj = oPages.AddItemTemplate("CMDBTN", "",  "", 0, 0, 15, 4, "ARIAL", 9,false ,false ,false ,false ,false , 1, "", "MODIFY",  "", False	)
NewObj.LabelOffsetY = -3
NewObj.LabelOffsetX = 0
NewObj.ItemBackColor = 26316
				
Set NewObj = oPages.AddItemTemplate("!TRIGGER", "",  "", 0, 0, 15, 4, "ARIAL", 9,false ,false ,false ,false ,false , 1, "", "MODIFY",  "", False	)
NewObj.LabelOffsetY = -3
NewObj.LabelOffsetX = 0
NewObj.ItemBackColor = 26316
			
Set NewObj = oPages.AddItemTemplate("!SUMMARY", "",  "", 0, 0, 15, 4, "ARIAL", 9,false ,false ,false ,false ,false , 1, "", "MODIFY",  "", False	)
NewObj.LabelOffsetY = -3
NewObj.LabelOffsetX = 0
NewObj.ItemBackColor = 26316
	
<% 
Do While Not oRS.EOF
	If oRS("CAPTION") = "-999999999" Then
		UseCaption = ReplaceStr(oRS("ATTRIBUTE_CAPTION"), """", """""")
	Else
		UseCaption = ReplaceStr(oRS("CAPTION"), """", """""")
	End If
	%>
	Set NewObj = oPages.AddItem ("<%= oRS("ATTR_INSTANCE_ID") %>", "<%= ReplaceStr(oRS("NAME"), """", """""")  %>",  "<%= UseCaption %>", <%= oRS("XPOS") %>, <%= oRS("YPOS") %>, <%= oRS("WIDTH") %>, <%= oRS("HEIGHT") %>, "ARIAL", 9,false ,false ,false ,false ,false , <%= Clng(oRS("SEQUENCE")) %>, "<%= oRS("INPUTTYPE") %>", "MODIFY",  "", False	)
	NewObj.readonly = true
	NewObj.MandatoryFlag = <%= ConBool(oRS("MANDATORY_FLG")) %>
	NewObj.SampleValue = "X"
	NewObj.SetExtraProperty "LUCOLUMN_NAME", true, "<%= ReplaceStr(oRS("LUCOLUMN_NAME"), """", """""") %>"
	NewObj.SetExtraProperty "LUDISPLAY_FLG", true, <%=ConBool( oRS("LUDISPLAY_FLG")) %>
	NewObj.SetExtraProperty "LUSTORAGE_FLG", true, <%= ConBool(oRS("LUSTORAGE_FLG")) %>
	NewObj.SetExtraProperty "LUSTORAGE_NAME", true, "<%= ReplaceStr(oRS("LUSTORAGE_NAME"), """", """""") %>"
	NewObj.SetExtraProperty "ATTRIBUTE_ID", true, "<%= oRS("ATTRIBUTE_ID") %>"
	NewObj.SetExtraProperty "TYPE" , true, "<%= oRS("TYPE") %>"
	NewObj.SetExtraProperty "ATTR_CAP", true, "<%= ReplaceStr(oRS("ATTRIBUTE_CAPTION"), """", """""") %>"
	NewObj.SetExtraProperty "CAPTION", true, "<%= ReplaceStr(oRS("CAPTION"), """", """""") %>"
	NewObj.SetExtraProperty "INPUTTYPE", true, "<%= ReplaceStr(oRS("INPUTTYPE"), """", """""") %>"
	NewObj.SetExtraProperty "ENTRYMASK", true, "<%= ReplaceStr(oRS("ENTRYMASK"), """", """""") %>"
	NewObj.SetExtraProperty "VALIDVALUEFIELD_FLG", true, "<%= ReplaceStr(oRS("VALIDVALUEFIELD_FLG"), """", """""") %>"
	NewObj.SetExtraProperty "reapplyoverride_flg", true, "<%= ConBool(oRS("REAPPLYOVERRIDE_FLG")) %>"
	NewObj.SetExtraProperty "DEFAULTVALUE", true, "<%= ReplaceStr(oRS("DEFAULTVALUE"), """", """""") %>"
	NewObj.SetExtraProperty "UNKNOWNVALUE", true, "<%= ReplaceStr(oRS("UNKNOWNVALUE"), """", """""") %>"
	NewObj.SetExtraProperty "TEXTLENGTH", true, "<%= ReplaceStr(oRS("TEXTLENGTH"), """", """""") %>"
	NewObj.SetExtraProperty "VISIBLERULE_ID", true, "<%= ReplaceStr(oRS("VISIBLERULE_ID"), """", """""") %>"
	NewObj.SetExtraProperty "ENABLEDRULE_ID", true, "<%= ReplaceStr(oRS("ENABLEDRULE_ID"), """", """""") %>"
	NewObj.SetExtraProperty "VALIDRULE_ID", true, "<%= ReplaceStr(oRS("VALIDRULE_ID"), """", """""") %>"
	NewObj.SetExtraProperty "PERSISTRULE_ID", true, "<%= ReplaceStr(oRS("PERSISTRULE_ID"), """", """""") %>"
	NewObj.SetExtraProperty "ACTION_ID", true, "<%= ReplaceStr(oRS("ACTION_ID"), """", """""") %>"
	NewObj.SetExtraProperty "SPELLCHECK_FLG", true, "<%= ReplaceStr(oRS("SPELLCHECK_FLG"), """", """""") %>"
	NewObj.SetExtraProperty "HELPSTRING", true, "<%= ReplaceStr(oRS("HELPSTRING"), """", """""") %>"
	NewObj.SetExtraProperty "DESCRIPTION", true, "<%= ReplaceStr(oRS("DESCRIPTION"), """", """""") %>"
	NewObj.SetExtraProperty "LU_TYPE_ID", true, "<%= ReplaceStr(oRS("LU_TYPE_ID"), """", """""") %>"
<% 
	If IsNull(oRS("ATTRIBUTEFRAME_ID")) Then
		ATTRIBUTEFRAME_ID = "null"
	Else
		ATTRIBUTEFRAME_ID = oRS("ATTRIBUTEFRAME_ID") 
	End If
%>
	NewObj.SetExtraProperty "ATTRIBUTEFRAME_ID", true, "<%= ReplaceStr(ATTRIBUTEFRAME_ID, """", """""") %>"
<%
	oRS.MoveNext
Loop	
oRS.close
Set oRS = Nothing
%>
LayoutCtl.SetContextMenu "DEFAULT", "COMMAND|Undo|145|COMMAND|" & G_Zoom_Text & "|100|COMMAND|" & G_ShowSelected_Text & "|125|COMMAND|" & G_ShowLabel_text & "|155|COMMAND|Save|110", "|"
LayoutCtl.SetContextMenu "FIELD", "COMMAND|Instance Properties...|400|COMMAND|Attribute Properties...|500|COMMAND|Over Ride|423|SEPARATOR|0|0|COMMAND|Remove|401", "|" 
LayoutCtl.CleanAllDirty
LayoutCtl.EnableTrackUndo = False
LayoutCtl.RedrawItems
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
-->
</SCRIPT>
<title>[ Frame id: <%=cFrameID%> ]</title>
</HEAD>
<BODY BGCOLOR=SILVER topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 CanDocUnloadNowInf=YES>
<DIV ID=HEADER style="position:absolute;width:102%;background-color:#006666;color:white;font:bold 12pt">&nbsp;<%= UCASE(oRS2("TITLE")) %></DIV><BR>
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
	<PARAM NAME="UseAnisotropicMapMode" VALUE="TRUE">
	<PARAM NAME="ShowGridBorder" VALUE="TRUE">
	<PARAM NAME="DefaultMandatoryBackColor" VALUE="255">
	<PARAM NAME="GridBorderX" VALUE="0">
	<PARAM NAME="GridBorderY" VALUE="0">
	<PARAM NAME="GridBorderWidth" VALUE=<%=Session("LayoutCtlWidth")%>>
	<PARAM NAME="GridBorderHeight" VALUE=<%=Session("LayoutCtlHeight")%>>
	<PARAM NAME="DefaultFontName" VALUE="Arial">
	<PARAM NAME="XViewExtent" VALUE="883">
	<PARAM NAME="YViewExtent" VALUE="533">
	<PARAM NAME="AutoCreateItem" VALUE="False">
</OBJECT>
</BODY>
</HTML>
<%
oRS2.Close
Set oRS2 = Nothing
oConn.Close
Set oConn = Nothing
%>
