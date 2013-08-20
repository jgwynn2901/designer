<!--#include file="..\lib\common.inc"-->
<%	Response.Expires = 0 
	MODE = CStr(Request.QueryString("MODE")) 
%>
<HTML>
<HEAD>
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
Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function
</SCRIPT>
<meta name="VI60_defaultClientScript" content="VBScript">
</HEAD>
   <frameset ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="SpDestSeqStepModal.asp?<%=Request.QueryString%>" scrolling="no" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>
</HTML>
