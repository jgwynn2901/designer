<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<TITLE>MyTabWindow</TITLE>
<script LANGUAGE="JavaScript">
function OnTabFramesReady()
{
	TheWindow.AddTab("General",120, "CF_General.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>");
	TheWindow.AddTab("Layout", 120, "CFLayout-f.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>");
	TheWindow.AddTab("Rules", 120, "CF_Rules.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>");
	TheWindow.AddTab("SQL Data", 120, "CF_SQLStatement.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>");
	TheWindow.AddTab("Listing", 120, "CF_ATTRIB_LIST.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>");
	TheWindow.AddTab("Frame Order", 120, "CF_FrameOrder.asp?FRAMEID=<%= Request.QueryString("FRAMEID") %>&CALLFLOW_ID=<%= Request.QueryString("CFID") %>");
	<% If Request.QueryString("ACTIVETAB") = "" Then %>
		TheWindow.SetActiveTab("General");
	<% Else %>
		TheWindow.SetActiveTab("<%= Request.QueryString("ACTIVETAB") %>");
	<% End If %>
}
</script>
</head>
<FRAMESET>
<FRAME FRAMEBORDER="0" NAME="TheWindow" SRC="CF_TabFrames.asp?CFID=<%= Request.QueryString("CFID") %>" COLS="*" WIDTH="1" HEIGHT="1">
</FRAMESET>
</html>
