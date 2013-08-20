<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetAHSID
	GetAHSID = frames("WORKAREA").GetAHSID
End Function

Function GetAHSIDName
	GetAHSIDName = frames("WORKAREA").GetAHSIDName
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

   <frameset framespacing="0" border="0" ROWS="190,*">
        <frame NAME="TOP" SRC="AHSSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="AHSSearchResults.asp" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>