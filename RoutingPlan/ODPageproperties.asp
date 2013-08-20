<!--#include file="..\lib\common.inc"-->
<% Response.Expires=0 %>
<HTML>
<HEAD>
<!--#include file="..\lib\tablecommon.inc"-->
<TITLE>Output Definition Field Properties</TITLE>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BTNCLOSE_onclick
	window.close
End Sub

Sub GetSelection (DYNKEY) 
	PROPVAL.innerHTML = DYNKEY
End Sub
-->
</SCRIPT>
<SCRIPT LANGUAGE="Javascript">
<!--
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	GetSelection( objRow.getAttribute("DYNKEY") )
}
-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Output Definition Field Properties</FONT>
</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back (1)" CLASS=LABEL>
U</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
S</TD>
</TR>
</TABLE>
</FIELDSET>

<DIV style="display:block;height:150;width:285;overflow:scroll">
<table cellPadding=2 cellSpacing=0 frame=void rules=all ID="tblFields" name="tblFields" width=100%  >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>Name:</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Value</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>(1) Name</td>
			<td NOWRAP CLASS=LABEL>Accord Auto</td>
		</tr>
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="2" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>(2) Background</td>
			<td NOWRAP CLASS=LABEL>AccordAuto.bmp</td>
		</tr>
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="3" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>(3) Output Tray</td>
			<td NOWRAP CLASS=LABEL>Top</td>
		</tr>
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="4" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>(4) Page Number</td>
			<td NOWRAP CLASS=LABEL>2</td>
		</tr>
	</tbody>
</table>
</DIV>
<TABLE>
<TR>
<TD CLASS=LABEL>Property:</TD>
<TD CLASS=LABEL><SPAN ID="PROPVAL"></SPAN></TD>
</TR>
<TR>
<TD CLASS=LABEL>Value:</TD>
<TD CLASS=LABEL><INPUT TYPE=TEXT CLASS=LABEL NAME="PROPTXTVALUE"></TD>
</TR>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BTNSAVE">Save</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BTNCLOSE">Close</BUTTON></TD>
</TR>
</TABLE>

</BODY>
</HTML>
