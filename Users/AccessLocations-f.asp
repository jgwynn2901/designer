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

Function GetACCID
	GetACCID = frames("WORKAREA").GetACCID
End Function

Function GetSelectedAHSID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedAHSID = document.all.tblFields.rows(idx).getAttribute("AHSID")
	Else
		GetSelectedAHSID = ""
	End If
End Function

Function GetAHSID

	GetAHSID = getmultipleindex(document.all.tblFields, "AHSID")
End Function

Function GetUID
	GetUID = frames("WORKAREA").GetUID
End Function
Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function

Function IsDirty
	IsDirty = frames("WORKAREA").CheckDirty
End Function
</script>

<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset CanDocUnloadNowInf="YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
        <frame NAME="WORKAREA" SRC="AccessLocation.asp?<%=Request.QueryString%>" SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
	</frameset>

</html>