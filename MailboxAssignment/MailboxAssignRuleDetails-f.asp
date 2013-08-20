<%
'***************************************************************
'Mailbox Assignment Rule Details frameset.
'
'$History: MailboxAssignRuleDetails-f.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:46p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MailboxAssignment
'* Hartford SRS: Initial revision
'***************************************************************
%>
<!--#include file="..\lib\common.inc"-->

<%	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	MODE = CStr(Request.QueryString("MODE"))
%>

<html>
<head>
<SCRIPT>
function CanDocUnloadNow()
{
<%if MODE <> "RO" then %>	
	b_IsRequired = frames("WORKAREA").f_CheckIsThisRequired();
	if (b_IsRequired == true)
	{
		alert("Each MAILBOX_ASSIGNMENT_TYPE must have at least 1 MAILBOX_ASSIGNMENT_RULE.");
		return false;
	}	
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
Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function
Function ExeCopy
	ExeCopy = frames("WORKAREA").ExeCopy
End Function

</SCRIPT>


<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset  CanDocUnloadNowInf= "YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="MailboxAssignRuleDetails.asp?<%=Request.QueryString%>"   SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>