<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
%>  
<html>
<head>
<base target="WORKAREA">
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>ZIP/Postal Code lookup</title>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=vbscript>
<!--

function Info_onsubmit
if document.Info.ZIP.value  = "" and document.Info.CITY.value  = "" and document.Info.STATE.value  = "" then
	msgbox "Please enter a Zip/Postal Code, a City or a State",vbExclamation
	Info_onsubmit=false
else
	Info_onsubmit=true
end if
end function

Sub window_onload()
<%
if Request.QueryString("ZIP") <> "" then
	%>
	document.Info.ZIP.value = "<%=UCase(Request.QueryString("ZIP"))%>"
	Info.submit 
	<%
end if
%>
End Sub

//-->
</SCRIPT>
</head>
<BODY  topmargin=20 leftmargin=10 rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» ZIP/Postal Code lookup Screen</td>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
</TR>
</TABLE></TD>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="70%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<form name="Info" action="ZIPLookupResults.asp" method="GET" target="WORKAREA">
<table>
<tr>
<TD CLASS=LABEL>Zip:<BR><INPUT TYPE=TEXT SIZE=7 NAME="ZIP" CLASS=LABEL></TD>
<TD CLASS=LABEL>City:<BR><INPUT TYPE=TEXT SIZE=23 NAME="CITY" CLASS=LABEL></TD>
<TD CLASS=LABEL>State/Province:<BR><INPUT TYPE=TEXT SIZE=2 NAME="STATE" CLASS=LABEL></TD>
<td width=100></td>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnLookup type=submit>Lookup</BUTTON></TD>
</tr>
</table>
</form>
</body>
</html>
