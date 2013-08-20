<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	MSFlexGrid.Font="Verdana"
	MSFlexGrid.AddItem  "Branch" , 1
	MSFlexGrid.AddItem  "Location" , 1
	MSFlexGrid.AddItem  "Assignment" , 1
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0 rightmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Branch Assignment Matrix</FONT>
</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
U</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
S</TD>
</TR>
</TABLE>
</FIELDSET>

<OBJECT CLASSID="clsid:6262D3A0-531B-11CF-91F6-C2863C385E30" 
	ID=MSFlexGrid
	WIDTH="100%"
	HEIGHT="75%">
<PARAM NAME=AllowUserResizing VALUE=1>
<PARAM NAME=Cols VALUE=6>
<PARAM NAME=Rows VALUE=15>
</OBJECT>


</BODY>
</HTML>
