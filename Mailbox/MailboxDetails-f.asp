<%
'***************************************************************
'Mailbox frameset.
'
'$History: MailboxDetails-f.asp $ 
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

Function GetMBID
	GetMBID = frames("WORKAREA").GetMBID
End Function

Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function

Function ExeCopy
	ExeCopy = frames("WORKAREA").ExeCopy
End Function

Function IsDirty
	IsDirty = frames("WORKAREA").CheckDirty
End Function
</SCRIPT>

<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset  CanDocUnloadNowInf= "YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="MailboxDetails.asp?<%=Request.QueryString%>"   scrolling=auto FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>