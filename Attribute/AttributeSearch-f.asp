<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetAID
	GetAID = frames("WORKAREA").GetAID
End Function

Function GetAIDName
	GetAIDName = frames("WORKAREA").GetAIDName
End Function

Function GetAIDCaption
	GetAIDCaption = frames("WORKAREA").GetAIDCaption
End Function

Function GetAIDInputType
	GetAIDInputType = frames("WORKAREA").GetAIDInputType
End Function

Function ExeSave
	MsgBox "Nothing to Save", 0 ,"FNSNetDesigner"
End Function

Function ExeCopy
	MsgBox "Nothing to Copy", 0 ,"FNSNetDesigner"
End Function

Function ExeDelete
	MsgBox "Nothing to Delete", 0 ,"FNSNetDesigner"
End Function

Function IsDirty
	IsDirty = false
End Function
</SCRIPT>
<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset  ROWS="180,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="AttributeSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="AttributeSearchResults.asp" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>