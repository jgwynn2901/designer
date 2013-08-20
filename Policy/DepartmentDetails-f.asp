<%
'***************************************************************
'DEPARTMENT frameset.
'
'$History: DepartmentDetails-f.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 1:39p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Policy
'* Added Department Interface due to ESIS Project.  It allows User to
'* create Department record attached to the AHSID in PROD Designer. The
'* permission used is the same as for Branch.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 12:10p
'* Created in $/FNS_DESIGNER/Source/Designer/Policy
'* Added Department Interface due to the ESIS Project.  It allows user to
'* attach AHSID to the department record.  Also, it allows user to delete,
'* create a new record and Edit an record in PROD Designer.  Permission
'* setup is the same as for Branch.  
'* 

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

Function GetDEPTID
	GetDEPTID = frames("WORKAREA").GetDEPTID
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
        <frame NAME="WORKAREA" SRC="DepartmentDetails.asp?<%=Request.QueryString%>"   scrolling=auto FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>