<HTML>
<HEAD>
<script LANGUAGE="JavaScript">

function ReCalcLayout()
{
		
	document.all.TOP.style.pixelTop = 0;
	document.all.TOP.style.pixelLeft = 0;
	document.all.TOP.style.pixelWidth = document.body.clientWidth;
	document.all.TOP.style.pixelHeight = 85;
	
	document.all.WORKAREA.top = document.all.TOP.style.pixelHeight + 1;
	document.all.WORKAREA.style.pixelLeft = 0;
	document.all.WORKAREA.style.pixelWidth = document.body.clientWidth;
	document.all.WORKAREA.style.pixelHeight = document.body.clientHeight - document.all.TOP.style.pixelHeight;

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
  <FRAMESET  ROWS="53,*,0,0" frameborder=no framespacing=0>
       <FRAME NAME="TOP" SRC="CFVisualEditor.asp?CFID=<%= Request.QueryString("CFID") %>&FRAMEID=<%= Request.querystring("FRAMEID") %>" SCROLLING=NO FRAMEBORDER="no" NORESIZE>
        <FRAME NAME="WORKAREA" SRC="ABOUT:BLANK" SCROLLING="auto" FRAMEBORDER="no">
        <FRAME NAME="HIDDENPAGE" SRC="ABOUT:BLANK" SCROLLING="no" FRAMEBORDER="no">
        <FRAME NAME="KEEPALIVE" SRC="..\Lib\KeepAlive.asp" SCROLLING="no" FRAMEBORDER="no">
	</FRAMESET>
</HEAD>
<BODY>
</BODY>
</HTML>




