<!--#include file="..\lib\common.inc"-->
<html>
<head>
<!--
'***************************************************************
'General purpose: Defines frameset for search and results frame
'$History: ClaimClassAssignRuleSearch-f.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/10/06    Time: 11:00p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/ClaimClass
'* New Claim Class Assignment module: Search, Details etc.

-->

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetCARID
	GetCARID = frames("WORKAREA").GetCARID
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
        <frame NAME="TOP" SRC="claimclassassignRuleSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="claimclassassignRuleSearchResults.asp" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>