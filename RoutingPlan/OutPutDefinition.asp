<!--#include file="..\lib\common.inc"-->
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<title>FNS Account Lookup Tree</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BtnEdit_onclick
	Parent.frames("WORK").location.href="OutputDefinitionEditor.asp"
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Output Definition Maintenance</FONT>
</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back (1)" CLASS=LABEL>
U</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
S</TD>
</TR>
</TABLE>
</FIELDSET>

<BR>
<FIELDSET>
<LEGEND CLASS=LABEL>Search</LEGEND>
<TABLE WIDTH="100%"><TR><TD>
<TABLE>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=25></TD>
<TD CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=25></TD>
</TR>
<TR>
<TD CLASS=LABEL>Contains output page with:<BR><INPUT TYPE=TEXT SIZE=25 CLASS=LABEL></TD>
<TD CLASS=LABEL></TD>
</TR>
</TABLE>
</TD>
<TD ALIGN=RIGHT>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON>Find</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON>Clear</BUTTON></TD>
</TR>
</TABLE>
</TD></TABLE>
<DIV style="display:block;height:120;width:450;overflow:scroll">
<table cellPadding=2 cellSpacing=0 frame=void rules=all ID="tblFields" name="tblFields" width=100%  >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>Name:</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Node</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>Desc</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>OD 1</td>
			<td NOWRAP CLASS=LABEL>Node</td>
			<td NOWRAP CLASS=LABEL>Desc Here</td>
		</tr>
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>OD 2</td>
			<td NOWRAP CLASS=LABEL>Node</td>
			<td NOWRAP CLASS=LABEL>Desc Here</td>
		</tr>
		<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>OD 3</td>
			<td NOWRAP CLASS=LABEL>Node</td>
			<td NOWRAP CLASS=LABEL>Desc Here</td>
		</tr>
				<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL>OD 3</td>
			<td NOWRAP CLASS=LABEL>Node</td>
			<td NOWRAP CLASS=LABEL>Desc Here</td>
		</tr>
	</tbody>
</table>
</DIV>
<BR>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnEdit">Edit</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON id=button1 name=button1>New</BUTTON></TD>
</TR>
</TABLE>
</FIELDSET>

</BODY>
</HTML>
