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
	TheWindow.AddTab("General",120, "OP_General.asp?OPID=<%= Request.QueryString("OPID") %>");
	TheWindow.AddTab("Layout", 120, "Layout-f.asp?OPID=<%= Request.QueryString("OPID") %>");
	TheWindow.AddTab("Listing", 120, "OP_Attribute_Listing.asp?OPID=<%= Request.QueryString("OPID") %>");
	<% If Request.QueryString("AHSID") <> "" Then %>
	TheWindow.AddTab("Override", 120, "Override-f.asp?<%= Request.QueryString %>");
	<% End If %>
	<% If Request.QueryString("ACTIVETAB") = "" Then %>
		TheWindow.SetActiveTab("General");
	<% Else %>
		TheWindow.SetActiveTab("<%= Request.QueryString("ACTIVETAB") %>");
	 <% End If %>
}
</script>
</head>
<FRAMESET>
<FRAME FRAMEBORDER="0" NAME="TheWindow" SRC="OP_TabFrames.asp?OPID=<%= Request.QueryString("OPID") %>" COLS="*" WIDTH="1" HEIGHT="1">
</FRAMESET>
</html>
