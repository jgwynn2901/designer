<!--#include file="..\lib\common.inc"-->

<%	Response.Expires = 0 

	MODE = CStr(Request.QueryString("MODE")) 
'***************************************************************
'General purpose: Defines frameset for claim class assignment
'
'$History: ClaimClassAssignRuleDetails-f.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/10/06    Time: 10:59p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/ClaimClass
'* New Claim Class Assignment module: Search, Details etc.



%>

<html>
<head>
<script>
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
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	frames("WORKAREA").PostTo(strURL)
End Sub

Function GetCARID
	GetCARID= frames("WORKAREA").GetCARID
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

</script>

<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset CanDocUnloadNowInf="YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="claimclassassignRuleDetails.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>