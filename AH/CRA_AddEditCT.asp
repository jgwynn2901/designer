<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnCancel_onclick()
{
window.close();
}

function BtnSave_onclick() 
{
if (document.frames("CCTFrame").document.readyState == "complete") 
	{
	document.frames("CCTFrame").document.frames("WORKAREA").ExeSave();
	}
}
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=12 topmargin=0>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Crawford Coverage Types </td>
</tr>
</table>

<iframe FRAMEBORDER="0" ID="CCTFrame" SRC="CRA_CTypes-f.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="90%">
</iframe>
<table cellpadding="2" cellspacing="1" ID="Table1">
<tr><nobr>
<td>
<BUTTON NAME=BtnSave ACCESSKEY="S" CLASS=STDBUTTON LANGUAGE=javascript onclick="return BtnSave_onclick()" ID="Button1"><U>S</U>ave</BUTTON></td>
<td>
<BUTTON NAME=BtnCancel ACCESSKEY="C" CLASS=STDBUTTON LANGUAGE=javascript onclick="return BtnCancel_onclick()" ID="Button2"><U>C</U>lose</BUTTON></td>
</tr>
</table>
</BODY>
</HTML>
