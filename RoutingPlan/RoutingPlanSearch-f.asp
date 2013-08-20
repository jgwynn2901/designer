<HTML>
<HEAD>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Function GetRPID
	GetRPID = frames("WORKAREA").GetRPID
End Function

Function GetRPDesc
	GetRPDesc = frames("WORKAREA").GetRPDesc
End Function

Function ExeSave
	MsgBox "Nothing to Save", 0 ,"FNSNetDesigner"
End Function

Function ExeCopy
	MsgBox "Nothing to Copy", 0 ,"FNSNetDesigner"
End Function

Function IsDirty
	IsDirty = false
End Function
</SCRIPT>
<meta name="VI60_defaultClientScript" content="VBScript">
</HEAD>

  <FRAMESET  ROWS="0,145, *" border=0 framespacing=0>
     	<FRAME NAME="hiddenPage" SRC=""  scrolling="No" noresize FRAMEBORDER="no" BORDER="0"  framespacing="0">
  		<FRAME NAME="TOP" SRC="RoutingPlanSearch.asp?CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>&ahsid=<%= Request.QueryString("ahsid") %>" SCROLLING=NO FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
  		<FRAME NAME="WORKAREA" SRC="RoutingPlanResults.asp" SCROLLING=NO FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</FRAMESET>
<BODY>
</BODY>
</HTML>
