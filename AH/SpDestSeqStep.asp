<!--#include file="..\lib\common.inc"-->
<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	MODE = "RW"
	if Request.QueryString("MODE") <> "" then MODE = Request.QueryString("MODE")
	dim bShowSave, bShowClose
	Select Case MODE
		Case "RO"
			bShowSave = false
			bShowClose = false
		Case "RW"
			bShowSave = true
			bShowClose = true
	End Select		
%>	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<Title>Specific Destination</Title>
<STYLE TYPE="text/css">HTML {width: 100pt; height: 100pt}</STYLE>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	document.frames("parentIFrame").document.frames.ExeSave();
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	if (document.frames("parentIFrame").CanDocUnloadNow() == true)
		window.close();
</script>

</HEAD>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="<%=BODYBGCOLOR%>">
<iFrame id="parentIFrame" Frameborder="0" src="SpDestSeqStep-f.asp?<%=Request.QueryString%>"  WIDTH="100%" Height="85%" ></iFrame>
<BR>
<TABLE>
	<TR>
			<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSave <% If MODE = "RO" Then Response.write(" DISABLED ") %> ACCESSKEY="S"><U>S</U>ave</BUTTON></TD>
			<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
	</TR>
</TABLE>
</BODY>
</HTML>