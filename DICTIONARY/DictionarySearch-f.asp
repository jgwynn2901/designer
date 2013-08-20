<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetDictText
	GetDictText = frames("WORKAREA").GetDictText
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
</script>

<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset ROWS="150,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="DictionarySearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="DictionarySearchResults.asp" SCROLLING="NO" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>

</html>