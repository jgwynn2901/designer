<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	document.all.TabFrame.style.height = document.body.clientHeight-20
	document.all.TabFrame.style.width = document.body.clientWidth
	document.frames("TabFrame").location.href = "SQLHELP-f.asp?SQL=FROM"
	window.setTimeout "ChngStatus", 5000
End Sub

Sub ChngStatus
	document.all.SPANSTATUS.style.display = "none"
End Sub

Sub BtnClose_onclick
	window.close()
End Sub

Sub BtnCopy_onclick
	Window.DialogArguments.name = document.frames("TabFrame").ExeCopy
End Sub

Sub BtnCopyColumn_onclick
If document.frames("TabFrame").document.frames("RIGHT").location.href <> "about:BLANK" Then
	Window.DialogArguments.name = document.frames("TabFrame").ExeCopyColumn
End If
End Sub
-->
</SCRIPT>
</HEAD>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1000" HEIGHT="1000"></iframe>
<CENTER>
<TABLE>
<TR>
<TD><SPAN ID=SPANSTATUS NAME=SPANSTATUS STYLE="DISPLAY:BLOCK;WIDTH:20;" CLASS=LABEL >Retrieving...</SPAN></TD>
<TD><BUTTON CLASS=STDBUTTON NAME=BtnCopy STYLE="WIDTH:110" ACCESSKEY="C"><U>C</U>opy Table Name</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON STYLE="WIDTH:110" NAME=BtnCopyColumn ACCESSKEY="P">Co<U>p</U>y Field Name</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON NAME=BtnClose ACCESSKEY="L">C<U>l</U>ose</BUTTON></TD>
</TR></TABLE>
</CENTER>
</BODY>
</HTML>
