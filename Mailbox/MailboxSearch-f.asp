<%
'***************************************************************
'Mailbox search frame.
'
'$History: MailboxSearch-f.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:46p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Mailbox
'* Hartford SRS: Initial revision
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("TOP").PostTo(strURL)
End Sub

Function GetMBID
	GetMBID = frames("WORKAREA").GetMBID
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
</head>
   <frameset  ROWS="140,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="MailboxSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="MailboxSearchResults.asp" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>