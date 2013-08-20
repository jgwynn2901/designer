<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
%>
<HTML>
<HEAD>
<STYLE TYPE="text/css">
HTML {width: 350pt; height: 140pt}
</STYLE>
<SCRIPT LANGUAGE=javascript>
var inObj;
function window.onload()
{
		inObj  = window.dialogArguments;
		
}
</SCRIPT>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnCancel_onclick() {
window.returnvalue = "cancel";
window.close()
}

function BtnSave_onclick() {
if (document.frames("AltFrame").document.frames("WORKAREA").document.all.ALT_NAME.value != "")
{
document.frames("AltFrame").document.frames("WORKAREA").ExeSave();
}
else
{
alert ("Alternate Name is a required field")
}
}
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Add Alternate Names </td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<iframe FRAMEBORDER="0" ID="AltFrame" SRC="Alt_Name-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="40%">
</iframe>
<BUTTON NAME=BtnSave ACCESSKEY="S" CLASS=STDBUTTON LANGUAGE=javascript onclick="return BtnSave_onclick()"><U>S</U>ave</BUTTON>
<BUTTON NAME=BtnCancel ACCESSKEY="C" CLASS=STDBUTTON LANGUAGE=javascript onclick="return BtnCancel_onclick()"><U>C</U>lose</BUTTON>
</BODY>
</HTML>
