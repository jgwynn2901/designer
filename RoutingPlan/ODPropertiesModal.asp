<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "ODPropertiesModal.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

%>
<HTML>
<HEAD>
<TITLE>Properties</TITLE>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT LANGUAGE=VBSCRIPT>
<!--
function PageCheck
StrError = ""
If Len(PropMAPPING.value) > 2550 Then
	StrError = StrError & "Mapping rule cannot be greater than 2550 character" & VbCrlf
End If

If Len(MYNAME.value) < 1 Then
	StrError = StrError & "Name is a required field" & VbCrlf
End if

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
		SampleValue.value = inObj.SampleValue;
		PropXPOS.value = inObj.xpos;
		PropYPOS.value = inObj.ypos;
		//PropITEMTYPE.value = inObj.itemtype;
		PropWIDTH.value = inObj.width;
		PropHEIGHT.value = inObj.height;
		FONT.value = inObj.fontname;
		PropSIZE.value = inObj.fontpointsize;
		PropBOLD.value = inObj.bold;
		PropITALIC.value = inObj.italic;
		PropUNDERLINE.value = inObj.underline;
		PropSTRIKEOUT.value = inObj.strikeout;
		PropMAPPING.value = inObj.formatstring;
		PropMULTILINE.value = inObj.multiline;
		MYNAME.value = inObj.name;
		BMP.value = inObj.bmp;
		inObj.pagestatus = "cancel";
}
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function BtnOK_onclick() 
{

		inObj.SampleValue = SampleValue.value;
		inObj.xpos = PropXPOS.value; 
		inObj.ypos = PropYPOS.value;
		//inObj.itemtype = PropITEMTYPE.value;
		inObj.width = PropWIDTH.value;
		inObj.height = PropHEIGHT.value;
		inObj.fontname = FONT.value;
		inObj.fontpointsize = PropSIZE.value;
		inObj.bold = PropBOLD.value;
		inObj.italic = PropITALIC.value;
		inObj.underline = PropUNDERLINE.value;
		inObj.strikeout = PropSTRIKEOUT.value;
		inObj.formatstring = PropMAPPING.value;
		inObj.multiline = PropMULTILINE.value;
		inObj.name = MYNAME.value;
		inObj.bmp = BMP.value;
		inObj.pagestatus = "save";		
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
//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub document_onkeydown
Select Case window.event.keyCode
	Case 13
		Call BtnOK_onclick()
	Case Else
End Select
End Sub
<!--#include file="..\lib\Help.asp"-->
-->
</SCRIPT>
</HEAD>
<BODY  BGCOLOR='<%=BODYBGCOLOR%>' >
<TABLE WIDTH=95% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Output Field Properties</SPAN>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=270 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<TABLE>
<TR>
<TD CLASS=LABEL COLSPAN=2>Name/Attribute:<BR>
<INPUT TYPE=TEXT CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> NAME=MYNAME SIZE=80 MAXLENGTH=80></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Sample Value:<BR>
<INPUT TYPE=TEXT CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> ID="SampleValue" NAME="SampleValue" SIZE=80 MAXLENGTH=80></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>XPos:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %>  NAME="PropXPOS" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>YPos:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> NAME="PropYPOS" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>Width:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> NAME="PropWIDTH" SIZE=8 MAXLENGTH=10></TD>
<TD CLASS=LABEL>Height:<BR>
<INPUT TYPE=TEXT  CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> NAME="PropHEIGHT" SIZE=8 MAXLENGTH=10></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Font:<BR>
<SELECT CLASS=LABEL  NAME=FONT <% If MODE="RO" Then Response.write(" DISABLED ") %>>
<OPTION value="Times New Roman">Times New Roman
<OPTION value="Courier New">Courier New
<OPTION value="Arial">Arial
</SELECT></TD>
<TD CLASS=LABEL>Font Size:<BR>
<SELECT CLASS=LABEL NAME="PropSIZE" STYLE="WIDTH:45" <% If MODE="RO" Then Response.write(" DISABLED ") %>>
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
</SELECT></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Bold:<BR>
 <SELECT CLASS=LABEL  NAME="PropBOLD" <% If MODE="RO" Then Response.write(" DISABLED ") %>>
<OPTION value="true">True
<OPTION value="false">False
</SELECT>
</TD>
<TD CLASS=LABEL>Italic:<BR>
 <SELECT CLASS=LABEL  NAME="PropITALIC" <% If MODE="RO" Then Response.write(" DISABLED ") %>>
<OPTION value="true">True
<OPTION value="false">False
</SELECT>
</TD>
<TD CLASS=LABEL>Underline:<BR>
 <SELECT CLASS=LABEL  NAME="PropUNDERLINE" <% If MODE="RO" Then Response.write(" DISABLED ") %>>
<OPTION value="true">True
<OPTION value="false">False
</SELECT>
</TD>
<TD CLASS=LABEL>Strikeout:<BR>
<SELECT  CLASS=LABEL  NAME="PropSTRIKEOUT" <% If MODE="RO" Then Response.write(" DISABLED ") %>>
<OPTION value="true">True
<OPTION value="false">False
</SELECT>
</TD>
<TD CLASS=LABEL>Multiline:<BR>
 <SELECT CLASS=LABEL  NAME="PropMULTILINE" <% If MODE="RO" Then Response.write(" DISABLED ") %>>
<OPTION value="true">True
<OPTION value="false">False
</SELECT>
</TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>BMP:<BR><INPUT TYPE=TEXT <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> CLASS=LABEL ID="BMP" NAME="BMP" SIZE=80 MAXLENGTH=255></TD>
</TR>

<TR>
<TD CLASS=LABEL>Mapping:<BR>
<TEXTAREA CLASS=LABEL  <% If MODE="RO" Then Response.write(" DISABLED ") %> NAME="PropMAPPING" STYLE="WIDTH:425;HEIGHT:100;"></TEXTAREA></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON <% If MODE="RO" Then Response.write(" DISABLED ") %> NAME=BtnOK LANGUAGE=javascript onclick="return BtnOK_onclick()">Ok</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel LANGUAGE=javascript onclick="return BtnCancel_onclick()">Cancel</BUTTON></TD>
</TR>
</TABLE>
</BODY>
</HTML>
