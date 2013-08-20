<% Response.expires=0 %>
<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</HEAD>
<BODY  leftmargin=0 topmargin=0 rightmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Agency Details</FONT>
</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
U</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
S</TD>
</TR>
</TABLE>
</FIELDSET>
<BR>
<LABEL CLASS=LABEL>Agency Details Go Here</LABEL>
<BR><BR><BR><BR>
<TABLE>
<TR>
<TD>
<TABLE BORDER=0>
<TR ROWSPAN=3>
<TD CLASS=LABEL>Agents:<BR>
<SELECT NAME="AGENTS" SIZE=4 STYLE="WIDTH:225" CLASS=LABEL>
<OPTION>Smith, Joe
<OPTION>Lynn, Amy
</SELECT>
</TD>
</TR>
</TABLE>
</TD>
<TD>
<TABLE>
<TR><TD><BUTTON CLASS=STDBUTTON>New</BUTTON></TD></TR>
<TR><TD><BUTTON CLASS=STDBUTTON id=button1 name=button1>Add</BUTTON></TD></TR>
<TR><TD><BUTTON CLASS=STDBUTTON id=button2 name=button2>Remove</BUTTON></TD></TR>
</TABLE>
</TD>
</TR>
</TABLE>
</BODY>
</HTML>
