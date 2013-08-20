<% Response.Expires=0 %>
<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<SCRIPT LANGUAGE="VBScript">
<!--
Function Dec(strHex)
    Dec = InStr("123456789ABCDEF", UCase(Left(strHex,1))) * 16
    Dec = Dec + InStr("123456789ABCDEF", UCase(Mid(strHex,2,1)))
End Function

Sub Document_OnKeyDown()
	if window.event.keyCode  = 13 or window.event.keyCode  = 8 then
		window.event.keyCode = 0
		window.event.returnValue=false
	end if
End Sub

Sub window_onLoad()
	Document.all("hhctrl").focus()
	Document.all("hhctrl").Click()
end sub
-->
</SCRIPT>
</HEAD>
<BODY topmargin=0 leftmargin=0  rightmargin=5 BGCOLOR="<%=BODYBGCOLOR%>">
<BR>
<OBJECT STYLE="Position:Absolute;LEFT:10;" id=hhctrl Name=hhctrl type="application/x-oleobject"
        classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11"
        codebase="../LIB/i386.cab#version=4,72,7291,0"
        width="98%"
        height="80%">
    <PARAM name="Command" value="Contents">
    <PARAM name="flags" value="0x0,0x35,0xFFFFFFFF">
    <PARAM name="Item1" value="Designer.hhc">
</OBJECT>
</BODY>
</HTML>