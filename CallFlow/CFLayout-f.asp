<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
lret = CheckDirty();
	if (true == lret)
	{
		if (false == confirm("Data has changed. Leave page without saving?"))
			return false;
		else
			return true;
	}
	else
	{
	return true;	
	}
}


function ReCalcLayout()
{
	document.all.STATUS.style.visibility = "hidden";
	
	document.all.TOPAREA.style.pixelTop = 0;
	document.all.TOPAREA.style.pixelLeft = 0;
	document.all.TOPAREA.style.pixelWidth = document.body.clientWidth;
	document.all.TOPAREA.style.pixelHeight = 45;
		
		
	document.all.LAYOUTAREA.top = document.all.TOPAREA.style.pixelHeight + 1;
	document.all.LAYOUTAREA.style.pixelLeft = 0;
	document.all.LAYOUTAREA.style.pixelWidth = document.body.clientWidth;
	document.all.LAYOUTAREA.style.pixelHeight = document.body.clientHeight - document.all.TOPAREA.style.pixelHeight;

}

function window.onresize()
{
	ReCalcLayout();
}

function window.onload()
{
	ReCalcLayout();
}
</script>
<SCRIPT LANGUAGE=VBScript>
<!--
Function CheckDirty()
		Set objCol =LAYOUTAREA.document.all.LayoutCtl.PageItems
	For Each x In objCol
		If x.dirty = True then
			CheckDirty = true
			Exit function
		End if
	Next
		CheckDirty = false
End Function
-->
</SCRIPT>
</HEAD>
<BODY  rightmargin=0 topmargin=0 leftmargin=0 BGCOLOR="#d6cfbd" CanDocUnloadNowInf=YES scroll=no>
<iframe FRAMEBORDER="0" SRC="hiddenPage.asp"  scrolling="No" noresize ID="hiddenPage" WIDTH="1" HEIGHT="1"></iframe>
<iframe FRAMEBORDER="0" SRC="CF_layoutTop.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>" SCROLLING=NO ID="TOPAREA" WIDTH="1" HEIGHT="1"></iframe>
<iframe FRAMEBORDER="0" SRC="CF_LayoutBottom.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>" SCROLLING="auto" ID="LAYOUTAREA" WIDTH="1" HEIGHT="1"></iframe>
<BR><SPAN CLASS=LABEL ID="STATUS">&nbsp Loading...</SPAN>
</BODY>
</HTML>
