<!--#include file="..\lib\common.inc"-->

<%	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	MODE = CStr(Request.QueryString("MODE"))
	ATID =  CStr(Request.QueryString("ATID"))
	ATLID =  CStr(Request.QueryString("ATLID"))		
%>
<html>
<head>
<script type="text/javascript" language="javascript">

function CanDocUnloadNow()
{
<%if MODE <> "RO" then %>	
	//b_IsRequired = frames("WORKAREA").f_CheckIsThisRequired();
	//if (b_IsRequired == true)
	//{
		//alert("Each Vendor Referral must have at least one Vendor Referral Rule.");
	//	return false;
	//}	
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
<%	End if %>
 
}
</script>

<script id="clientEventHandlersVBS" language="vbscript" type="text/vbscript">
Function ExeSave
	ExeSave = frames("WORKAREA").ExeSave
End Function
</script>


<meta name="VI60_defaultClientScript" content="VBScript">
</head>
   <frameset  CanDocUnloadNowInf= "YES" ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize frameborder="no" BORDER="0" framespacing="0"/>
        <frame NAME="WORKAREA" SRC="AccountTipDetails.asp?<%=Request.QueryString%>"   scrolling="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0"/>
	</frameset>

</html>