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

Function GetAID
	GetAID = frames("WORKAREA").GetAID
End Function

Function GetAIDName
	GetAIDName = frames("WORKAREA").GetAIDName
End Function

Function GetAIDCaption
	GetAIDCaption = frames("WORKAREA").GetAIDCaption
End Function

Function GetAIDInputType
	GetAIDInputType = frames("WORKAREA").GetAIDInputType
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
        <frame NAME="WORKAREA" SRC="AttributeDetails.asp?<%=Request.QueryString%>"   scrolling=auto FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>