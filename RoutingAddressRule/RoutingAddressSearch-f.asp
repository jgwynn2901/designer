<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetRAID
	GetRAID = frames("WORKAREA").GetRAID
End Function

Function GetRAIDDescription 
	GetRAIDDescription = frames("WORKAREA").GetRAIDDescription
End Function

Function GetRAIDState
	GetRAIDState = frames("WORKAREA").GetRAIDState
End Function

Function GetRAIDFIPS
	GetRAIDFIPS = frames("WORKAREA").GetRAIDFIPS
End Function

Function GetRAIDZip
	GetRAIDZip = frames("WORKAREA").GetRAIDZip
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
   <frameset ROWS="180,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="RoutingAddressSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="RoutingAddressSearchResults.asp" SCROLLING="NO" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>

</html> 