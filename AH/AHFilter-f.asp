<!--#include file="..\lib\common.inc"-->

<script>
function AddNode(AHSID,NAME,WHICHLIST)
{
	frames("NODEFILTER").AddNode(AHSID,NAME,WHICHLIST);
}

function ExeClear()
{
	frames("NODEFILTER").ExeClear();
	frames("WORKAREA").ExeClear();
	frames("TOP").ClearSearch();
}

</script>
<HTML>
<HEAD>
   <FRAMESET  ROWS="0,140,*,100" border=0 framespacing=0>
   		<FRAME NAME="hiddenPage" SRC="AHFilterModify.asp"  scrolling="No" noresize FRAMEBORDER="no" BORDER="0"  framespacing="0">
        <FRAME NAME="TOP" SRC="AHFilter.asp?AHSID=<%=Request.QueryString("AHSID")%>" SCROLLING=NO FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <FRAME NAME="WORKAREA" SRC="AHFilterSearchResults.asp?AHSID=<%=Request.QueryString("AHSID")%>" SCROLLING="auto" FRAMEBORDER="no" BORDER="0" framespacing="0">
        <FRAME NAME="NODEFILTER" SRC="AHFilterNode.asp?AHSID=<%=Request.QueryString("AHSID")%>" SCROLLING=NO FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
     </FRAMESET>
</HEAD>
<BODY>
</BODY>
</HTML>