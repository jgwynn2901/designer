<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetRID
	GetRID = frames("WORKAREA").GetRID
End Function

Function GetRIDText
	GetRIDText = frames("WORKAREA").GetRIDText
End Function

Function GetRIDType
	GetRIDType = frames("WORKAREA").GetRIDType
End Function

Function GetRIDComments
	GetRIDComments = frames("WORKAREA").GetRIDComments
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
   <frameset ROWS="190,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="RuleSearch.asp?<%=Request.QueryString%>" SCROLLING="no" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="RuleSearchResults.asp" SCROLLING="NO" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>

</html>