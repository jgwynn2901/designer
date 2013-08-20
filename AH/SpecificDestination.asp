<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	MODE = "RW"
	
	if CStr(Request.QueryString("MODE")) <> "" then	MODE = CStr(Request.QueryString("MODE"))
	
	DIM bShowSave, bShowClose
	Select Case MODE
		Case "RO"
			bShowSave = false
			bShowClose = false
		Case "RW"
			bShowSave = TRUE
			bShowClose = TRUE
	End Select
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>Specific Destination</Title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<STYLE TYPE="text/css">HTML {width: 350pt; height: 700pt}</STYLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onunload() {
	if (window.frames("parentIFrame").window.frames("WORKAREA").span_SDID.innerHTML != "NEW")
		window.dialogArguments.SDID = window.frames("parentIFrame").window.frames("WORKAREA").span_SDID.innerHTML;
}
//-->
</SCRIPT>
<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	//document.frames("parentIFrame").document.frames.ExeSave();
	if (document.frames("parentIFrame").document.frames.ExeSave() == true)
		document.all.BtnSave.disabled = true
</SCRIPT>
<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("parentIFrame").CanDocUnloadNow() == true)
		window.close();
</script>
</HEAD>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" LANGUAGE=javascript onunload="return window_onunload()">
<iFrame id="parentIFrame" Frameborder="0" src="specificdestination-f.asp?<%=Request.QueryString%>" Height="90%" Width="100%"></iFrame>
<BR>
<TABLE>
	<TR>
	<% If Request.QueryString("SDID") = "NEW" Then %>
		<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSave <% If MODE = "RO" Then Response.write(" DISABLED ") %> ACCESSKEY="S"><U>S</U>ave</BUTTON></TD>
	<% End If %>
		<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
	</TR>
</TABLE>
</BODY>
</HTML>