<%
'***************************************************************
'Mailbox frameset.
'
'$History: MyGreetingDetails-f.asp $ 
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
'* User: Alex.shimberg Date: 4/30/06    Time: 9:45p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Mailbox
'* Hartford SRS: Initial revision
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->

<%	Response.Expires = 0 

	MODE = CStr(Request.QueryString("MODE")) 
%>

<html>
<head>
<SCRIPT>
function CanDocUnloadNow()
{
<%	if MODE <> "RO" then %>	
	bDirty = frames("WORKAREA").CheckDirty();
	
	if (bDirty == true)
	{
		if (false == confirm("Data has changed. Leave page without saving?"))
			return false;
		else
			return true;
	}
	else
		return true;
		
<%	else %>
	return true;
<%	end if %>
 
}
</SCRIPT>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("WORKAREA").PostTo(strURL)
End Sub

Function GetGreetingID
	GetGreetingID = frames("WORKAREA").GetGreetingID
End Function

Function GetGreetingText
	GetGreetingText = frames("WORKAREA").GetGreetingText
End Function

Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function

Function ExeCopy
	ExeCopy = frames("WORKAREA").ExeCopy
End Function

Function ExeDelete
	ExeDelete = frames("WORKAREA").ExeDelete
End Function

Function IsDirty
	IsDirty = frames("WORKAREA").CheckDirty
End Function
</SCRIPT>

<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset  CanDocUnloadNowInf= "YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="MyGreetingDetails.asp?<%=Request.QueryString%>"   scrolling=auto FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>