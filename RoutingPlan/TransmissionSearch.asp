<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BtnSelect_onclick
	Parent.frames("BOTTOM").location.href = "TransmissionDetails.asp?STATUS=UPDATE&TRANSMISSION_TYPE_ID="& document.all.TRANSMISSION_TYPE_ID.value
End Sub

Sub BtnNew_onclick
	Parent.frames("BOTTOM").location.href = "TransmissionDetails.asp?STATUS=NEW&TRANSMISSION_TYPE_ID=NEW"
End Sub

-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR=#d6cfbd topmargin=5 rightmargin=0 leftmargin=0 CanDocUnloadNowInf=NO>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Transmission Types
</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<TABLE WIDTH="100%">
<TR><TD VALIGN=TOP>
<TABLE>
<TR>
<TD CLASS=LABEL>Transmission Types:<BR>
<SELECT NAME="TRANSMISSION_TYPE_ID" CLASS=LABEL>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQL = ""
	SQL = SQL & "SELECT * FROM TRANSMISSION_TYPE"
	Set RS = conn.Execute(SQL)
	Do While Not RS.EOF
%>
<OPTION VALUE="<%= RS("TRANSMISSION_TYPE_ID") %>"><%= RS("NAME") %>
<%
RS.MoveNext
Loop
RS.Close
%>
</SELECT>
</TD>
</TR>
</TABLE>
</TD><TD ALIGN=RIGHT>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelect ACCESSKEY="S"><U>S</U>elect</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnNew  ACCESSKEY="N"><U>N</U>ew</BUTTON></TD>
</TR>
</TABLE>
</TD></TR>
</TABLE>
</BODY>
</HTML>
