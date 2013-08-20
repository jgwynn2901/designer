<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetBID
	GetBID = frames("WORKAREA").GetBID
End Function

Function GetCIDName
	GetCIDName = frames("WORKAREA").GetCIDName
End Function

Function GetCIDCaption
	GetCIDCaption = frames("WORKAREA").GetCIDCaption
End Function

Function GetCIDInputType
	GetCIDInputType = frames("WORKAREA").GetCIDInputType
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

</head>
   <frameset  ROWS="155,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="BillingSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="BillingSearchResults.asp" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>