<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<style TYPE="text/css">
HTML {width: 300pt; height:125pt}
</style>

<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Confirmation Dialog</TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BtnDelete_onclick
	window.returnvalue = "DELETE"
	window.close
End Sub


Sub BtnCancel_onclick
	window.returnvalue = ""
	window.close
End Sub

-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR=#d6cfbd topmargin=0 rightmargin=0 leftmargin=0  bottommargin=0>

<SPAN CLASS=LABEL STYLE="COLOR:#FF0000;FONT-SIZE:11pt">
Are you sure you want to delete this call flow?<BR>
<% If Clng(Request.QueryString("COUNT")) > 1 Then %>
This call flow is shared by other accounts!<BR>
<% End If %>
This action can not be undone!
</SPAN>
<BR><BR><BR>
<CENTER>
<TABLE>
<TR>
<TD>
<BUTTON CLASS=STDBUTTON NAME=BtnDelete ACCESSKEY="D"><U>D</U>elete</BUTTON>
</TD>
<TD>
<BUTTON CLASS=STDBUTTON NAME=BtnCancel ACCESSKEY="C"><U>C</U>ancel</BUTTON>
</TD>
</TR>
</TABLE>
</CENTER>
</BODY>
</HTML>
