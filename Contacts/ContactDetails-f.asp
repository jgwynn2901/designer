<!--#include file="..\lib\common.inc"-->

<%	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
%>

<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function


</SCRIPT>


<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="ContactDetails.asp?<%=Request.QueryString%>"   SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>