<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetOPID
	GetOPID = frames("WORKAREA").GetOPID
End Function

Function GetOPIDName
	GetOPIDName = frames("WORKAREA").GetOPIDName
End Function

Function GetOPIDCaption
	GetOPIDCaption = frames("WORKAREA").GetOPIDCaption
End Function

Function GetOPIDInputType
	GetOPIDInputType = frames("WORKAREA").GetOPIDInputType
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
</head>

   <frameset  ROWS="150,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="OutputPageSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="OutputPageSearchResults.asp" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>