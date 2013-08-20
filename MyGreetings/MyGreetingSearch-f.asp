<%
'***************************************************************
'Mailbox search frame.
'
'$History: MyGreetingSearch-f.asp $ 
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:59p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:35p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:33p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:14p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:09p
'* Created in $/FNS_DESIGNER/Source/Designer/Greeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 4/21/08    Time: 9:23a
'* Created in $/FNS_DESIGNER/Source/Designer
'* created for Sedgwick.  Just want to save my work for now
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

Function GetGreetingID
	GetGreetingID = frames("WORKAREA").GetGreetingID
End Function

Function GetGreetingText
	GetGreetingText = frames("WORKAREA").GetGreetingText
End Function

Function ExeSave
	MsgBox "Nothing to Save", 0 ,"FNSNetDesigner"
End Function

Function ExeCopy
	MsgBox "Nothing to Copy", 0 ,"FNSNetDesigner"
End Function

Function ExeDelete
	ExeDelete = frames("WORKAREA").ExeDelete
End Function

Function IsDirty
	IsDirty = false
End Function
</SCRIPT>
<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset  ROWS="140,*" border="0" framespacing="0">
        <frame NAME="TOP" SRC="MyGreetingSearch.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="MyGreetingSearchResults.asp" scrolling= "NO" FRAMEBORDER="no" BORDER="0" framespacing="0">
	</frameset>
</html>