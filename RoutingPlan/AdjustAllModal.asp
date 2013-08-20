<!--#include file="..\lib\common.inc"-->
<% Response.expires = 0 %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	window.dialogArguments.pagestatus = "cancel"
End Sub

Sub BtnHelp_onclick
	strURL = "HTTP://<%= Request.Servervariables("server_name") %>/FNSdesigner/Help/AdjustAllOD.Html"
	lret = window.showHelp(strURL)
End Sub

Sub BtnCancel_onclick
	window.close()
End Sub

Sub BtnAdjust_onclick
strerror = ""
If ADJUSTX.value = "" AND ADJUSTY.value = "" AND ADJUSTWIDTH.value = "" AND ADJUSTHEIGHT.value = "" Then
	strerror = strerror & "At least one adjustment must be entered" & vbcrlf
End If 
If ADJUSTX.value <> "" AND Not Isnumeric(ADJUSTX.value) Then
	strerror = strerror & "X Adjustment must be numeric" & vbcrlf
End If
If ADJUSTY.value <> "" AND Not Isnumeric(ADJUSTY.value) Then
	strerror = strerror & "Y Adjustment must be numeric" & vbcrlf
End If
If ADJUSTWIDTH.value <> "" AND Not Isnumeric(ADJUSTWIDTH.value) Then
	strerror = strerror & "Width Adjustment must be numeric" & vbcrlf
End If
If ADJUSTHEIGHT.value <> "" AND Not Isnumeric(ADJUSTHEIGHT.value) Then
	strerror = strerror & "Height Adjustment must be numeric" & vbcrlf
End If
if strerror  = "" Then
	window.dialogArguments.Adjust_X = ADJUSTX.value 
	window.dialogArguments.Adjust_Y = ADJUSTY.value
	window.dialogArguments.Adjust_Height = ADJUSTWIDTH.value
	window.dialogArguments.Adjust_Width = ADJUSTHEIGHT.value
	window.close()
else
	msgbox strerror,0,"FNSDesigner"
End if 

End Sub

-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR="#d6cfbd">
<TABLE WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Adjust All:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help ID=BtnHelp NAME=BtnHelp></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=100 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=100 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
<TR>
<TD CLASS=LABEL>X-Coordinate(+/-):<BR><INPUT TYPE=TEXT NAME=ADJUSTX CLASS=LABEL></TD>
<TD CLASS=LABEL>Y-Coordinate(+/-):<BR><INPUT TYPE=TEXT NAME=ADJUSTY CLASS=LABEL></TD>
</TR>
<TR>
<TD CLASS=LABEL>Width(+/-):<BR><INPUT TYPE=TEXT NAME=ADJUSTWIDTH CLASS=LABEL></TD>
<TD CLASS=LABEL>Height(+/-):<BR><INPUT TYPE=TEXT NAME=ADJUSTHEIGHT CLASS=LABEL></TD>
</TR>
</TABLE>
<BR>
<TABLE>
<TR>
<TD><BUTTON NAME=BtnAdjust ID=BtnAdjust CLASS=STDBUTTON ACCESSKEY="A">Adjust</BUTTON></TD>
<TD><BUTTON NAME=BtnCancel ID=BtnCancel CLASS=STDBUTTON ACCESSKEY="C">Cancel</BUTTON></TD>
</TR>
</TABLE>
</BODY>
</HTML>
