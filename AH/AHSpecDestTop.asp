<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<%

cAHID = Request.QueryString("AHID")
cName = Request.QueryString("NAME")
cCN = Request.QueryString("CLIENT_NODE")
if trim(cCN) <> "" then
	cCustID = cCN
else
	cCustID = cAHID
end if	

%>
<SCRIPT language="vbScript">
sub BtnClose_onclick()
window.close
end sub
</SCRIPT>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Specific Destination</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="JavaScript">
<!--
function SpecificDestObj()
{
	this.specDest = "";
}

var oSD = new SpecificDestObj();
// -->
</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	oSD = window.dialogArguments;
	window.defaultStatus = "Ready";
</script>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onunload
dim x, cSD, lFirst, oTable

lFirst = true
set oTable = document.frames(0).document.all.tblResult
for x=0 to oTable.rows.length-1
	if oTable.rows(x).className = "ResultRow" then
		if lFirst then
			cSD = oTable.rows(x).getAttribute("SDID")
			lFirst = false
		else
			cSD = cSD & " " & oTable.rows(x).getAttribute("SDID")		
		end if
	end if
next
oSD.specDest = cSD
End Sub

-->
</SCRIPT>
</head>
<body leftmargin="15" topmargin="0" BGCOLOR="#d6cfbd" bottommargin="0" rightmargin="0">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Specific Destination(s)</td>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table CELLSPACING="0" CELLPADDING="0" WIDTH="300" BORDER="0" STYLE="BACKGROUND-COLOR:Seashell">
<tr>
<td CLASS="LABEL"><b>Name: </b> <%= cName %></td>
</tr>
<tr>
<td CLASS="LABEL"><b>ID: </b><%= cAHID %></td>
</tr>
<tr><td CLASS="LABEL"><br></td></tr>
</table>
<IFRAME id="f1" NAME="WORKAREA" FRAMEBORDER="0" align="bottom" SRC="AHSpecDestBott.asp?SD=<%=Request.QueryString("SD")%>&AHSID=<%=cAHID%>&NAME=<%=server.URLEncode(cName)%>&CLIENT_NODE=<%=cCustID%>&MODE=<%=Request.QueryString("MODE")%>" WIDTH="100%" HEIGHT="73%"></IFRAME>
<table align="left" width="100%">
<tr>
<td CLASS="LABEL" align="left"><button CLASS="StdButton" NAME="BtnClose">Close</button></td>
</tr>
</table>
</body>
</html>

