<!--#include file="..\lib\common.inc"-->

<%	Response.Expires = 0 
	MODE = CStr(Request.QueryString("MODE")) 
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

Function GetPID
	GetPID = frames("WORKAREA").GetPID
End Function
Function GetLID
	GetLID = frames("WORKAREA").GetLID
End Function


Function ExeSave
	MsgBox "Nothing to Save", 0 ,"FNSNetDesigner"
End Function


Function IsDirty
	IsDirty = frames("WORKAREA").CheckDirty
End Function

</script>

<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset CanDocUnloadNowInf="YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="LocationDetails.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>