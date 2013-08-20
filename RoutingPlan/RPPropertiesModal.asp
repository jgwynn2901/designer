<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Transmission Type Properties</TITLE>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--


-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=5 topmargin=0>
<BR>
<TABLE CELLPADDING=2 CELLSPACING=0>
<TR>
<TD CLASS=LABEL COLSPAN=2>Transmission Type:<BR>
<SELECT NAME="TTYPE" CLASS=LABEL STYLE="WIDTH:150">
<OPTION VALUE="Fax">Fax
<OPTION VALUE="Print">Print
<OPTION VALUE="ICMSRouting">ICMSRouting
</SELECT>
</TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Destination String:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="DSTRING" STYLE="WIDTH:150"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Retry Wait:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="WAIT" STYLE="WIDTH:75"></TD>
<TD CLASS=LABEL>Retry Count:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="COUNT" STYLE="WIDTH:74"></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Sequence:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="SEQUENCE" SIZE=2>of 2</TD>
</TR>
</TABLE>
</BODY>
</HTML>
