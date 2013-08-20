<!--#include file="..\Lib\Common.inc"-->
<%
response.expires=0
%>
<script LANGUAGE="VBScript">
<!--
Sub Document_OnKeyDown()
	if window.event.keyCode  = 13 or window.event.keyCode  = 8 then
		window.event.keyCode = 0
		window.event.returnValue=false
	end if
End Sub

Sub window_onLoad()
end sub
-->
</script>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</HEAD>
<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>"><font face="ms sans serif">
<FORM NAME=FrmHelp>
<TABLE>
<TR>
<TD CLASS=LABEL>
Type the word(s) you are looking for, then click the index entry you want and click the Display Button.
</TD>
</TR>
</TABLE>
<OBJECT id=hhctrl type="application/x-oleobject"
	  		 classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11"
			 codebase="../Lib/i386.cab#version=4,72,7291,0"
			 width=100%
			 height=220>
<PARAM name="Command" value="index">
<PARAM name="flags" value="0x0,0x35,0xFFFFFFFF">
<PARAM name="Item1" value="Designer.hhk">
</OBJECT>
</form>
</BODY>
</HTML>