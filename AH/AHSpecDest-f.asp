<!--#include file="..\lib\common.inc"-->
<%
dim cAHID, cName

cAHID = Request.QueryString("AHID")
cName = Request.QueryString("NAME")
%>
<html>
<head>
</head>
	<IFRAME ID=FRTOP FRAMEBORDER=0 SCROLLING=NO SRC="AHSpecDestTop.asp?AHID=<%=cAHID%>&NAME=<%=cName%>"></IFRAME>
	<IFRAME ID=FRBOT FRAMEBORDER=0 SCROLLING=AUTO SRC="AHSpecDestBott.asp?SD=<%=Request.QueryString("SD")%>"></IFRAME>
</html>