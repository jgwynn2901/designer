<!--#include file="..\lib\common.inc"-->
<html>
<HEAD>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function ExeCopy
	ExeCopy = frames("LEFT").GetTableName()
End Function

Function ExeCopyColumn
	ExeCopyColumn = frames("RIGHT").GetColumnName()
End Function
-->
</SCRIPT>
</HEAD>
   <frameset  COLS="300,*" border="0" framespacing="0">
        <frame NAME="LEFT" SRC="SQLHELP.asp?<%=Request.QueryString%>" SCROLLING="AUTO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="RIGHT" SRC="ABOUT:BLANK" SCROLLING="AUTO" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>

