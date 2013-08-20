<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetSDID
	GetCOID = frames("WORKAREA").GetSDID
End Function

</SCRIPT>
</head>
<FRAMESET  ROWS="190,*" border=0 framespacing=0>
  	<FRAME NAME="TOP" SRC="AHSpecDestSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
  	<FRAME NAME="WORKAREA" SRC="AHSpecDestResults.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" BORDER="0" framespacing="0">
</FRAMESET>
</HTML>
