<%
Response.Expires = 0
%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function RefreshMe()
	self.location.href = "KeepAlive.asp"
End Function

Sub window_onload
'reload every 15 minutes
	lret = Window.setInterval("RefreshMe()",   900000)
End Sub
-->
</SCRIPT>
</HEAD>
<BODY>
<%= Now() %>
</BODY>
</HTML>
