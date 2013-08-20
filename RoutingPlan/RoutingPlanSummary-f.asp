<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
   <FRAMESET  ROWS="0,190,*" border=0 framespacing=0>
   		<FRAME NAME="hiddenPage" SRC="ABOUT:BLANK"  scrolling="No" noresize FRAMEBORDER="no" BORDER="0"  framespacing="0">
        <FRAME NAME="TOP" SRC="RoutingPlanSummary.asp?ROUTING_PLAN_ID=<%= Request.QueryString("ROUTING_PLAN_ID") %>&AHSID=<%= Request.QueryString("AHSID") %>" SCROLLING=NO FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
        <FRAME NAME="WORKAREA" SRC="RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.QueryString("ROUTING_PLAN_ID") %>&AHSID=<%= Request.QueryString("AHSID") %>" SCROLLING=NO FRAMEBORDER="no" BORDER="0" framespacing="0">
	</FRAMESET>
</HEAD>
<BODY>
</BODY>
</HTML>
