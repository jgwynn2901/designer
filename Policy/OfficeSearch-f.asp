<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetOID
	GetOID = frames("WORKAREA").GetOID
End Function

Function GetOIDType
	GetOIDType = frames("WORKAREA").GetOIDType
End Function

Function GetOIDState
	GetOIDState = frames("WORKAREA").GetOIDState
End Function

Function GetOIDNumber
	GetOIDNumber = frames("WORKAREA").GetOIDNumber
End Function

Function GetOIDZip
	GetOIDZip = frames("WORKAREA").GetOIDZip
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
        <frame NAME="TOP" SRC="OfficeSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="OfficeSearchResults.asp" SCROLLING="NO" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>

</html> 