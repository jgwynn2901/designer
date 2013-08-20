<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT LANGUAGE=VBScript>
<!--
function CheckOptions()
strError = ""
if Len(minimumheight.value) < 1 Then
	strError = strError & "Minimum height is a required field" & VbCrlf
End if

if Len(defaultheight.value) < 1 OR Not Isnumeric(defaultheight.value) Then
	strError = strError & "Default height is a required field, and must be numeric" & VbCrlf
End if

if Len(defaultwidth.value) < 1 OR Not Isnumeric(defaultwidth.value) Then
	strError = strError & "Default width is a required field, and must be numeric" & VbCrlf
End if

if Len(minimumwidth.value) < 1 OR Not Isnumeric(minimumwidth.value) Then
	strError = strError & "Minimum width is a required field, and must be numeric" & VbCrlf
End if

If strError <> "" Then
	MsgBox strError, 0 , "FNSDesigner"
	CheckOptions = false
else
	CheckOptions = true
end if
end function
-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
var OptionsObj;
function window.onload()
{
	OptionsObj  = window.dialogArguments;
	//alert (OptionsObj.defaultfont + "|" + OptionsObj.defaultfontsize )
	defaultheight.value  = OptionsObj.defaultheight;
	defaultwidth.value = OptionsObj.defaultwidth;
	defaulfontsize.value = OptionsObj.defaultfontsize;
	defaultfont.value = OptionsObj.defaultfont;
	minimumwidth.value = OptionsObj.minimumwidth;
	minimumheight.value = OptionsObj.minimumheight;
	OptionsObj.pagestatus = "cancel";
}

</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function BtnCancel_onclick() {
OptionsObj.pagestatus = "cancel";
window.returnvalue = null;
window.close();
}

function BtnOK_onclick() {
	OptionsObj.defaultheight = defaultheight.value;
	OptionsObj.defaultwidth = defaultwidth.value;
	OptionsObj.defaultfontsize = defaulfontsize.value;
	OptionsObj.defaultfont = defaultfont.value;
	OptionsObj.minimumwidth = minimumwidth.value;
	OptionsObj.minimumheight = minimumheight.value;
	OptionsObj.pagestatus = "save";
	lret = CheckOptions()
	
	if (true == lret)
	{
	window.returnvalue = OptionsObj;
	window.close();
	}
}

//-->
</SCRIPT>
<TITLE>Default Layout Control Options</TITLE>
</HEAD>

<BODY BGCOLOR="#d6cfbd">
<TABLE WIDTH=200 CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Default Options</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=200 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=200 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<TABLE>
<TR>
<TD CLASS=LABEL>Default Font:<BR>
<SELECT NAME="defaultfont" CLASS=LABEL STYLE="WIDTH:140">
<OPTION VALUE="Times New Roman">Times New Roman
<OPTION value="Courier New">Courier New
<OPTION VALUE="Arial">Arial
</SELECT>
</TD>
<TD CLASS=LABEL>Default Font Size:<BR>
<SELECT NAME="defaulfontsize" CLASS=LABEL>
<OPTION VALUE="6">6
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
</TR>
<TR>
<TD CLASS=LABEL>Default Width:<BR><INPUT TYPE=TEXT MAXLENGTH=10 NAME=defaultwidth SIZE=25 CLASS=LABEL></TD>
<TD CLASS=LABEL>Default Height:<BR><INPUT TYPE=TEXT MAXLENGTH=10  NAME=defaultheight SIZE=25 CLASS=LABEL></TD>
</TR>
<TR>
<TD CLASS=LABEL>Minimum Width:<BR><INPUT TYPE=TEXT MAXLENGTH=10  NAME=minimumwidth SIZE=25 CLASS=LABEL></TD>
<TD CLASS=LABEL>Minimum Height:<BR><INPUT TYPE=TEXT MAXLENGTH=10  NAME=minimumheight SIZE=25 CLASS=LABEL></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD><BUTTON CLASS=STDBUTTON NAME=BtnOK LANGUAGE=javascript onclick="return BtnOK_onclick()">Ok</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON NAME=BtnCancel LANGUAGE=javascript onclick="return BtnCancel_onclick()">Cancel</BUTTON></TD>
</TR>
</TABLE>
</BODY>
</HTML>
