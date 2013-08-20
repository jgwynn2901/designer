
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Network Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="#d6cfbd">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Network Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<form Name="FrmSearch" METHOD="GET" ACTION="CarrierSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="RW">
<input type="hidden" NAME="CID" value="">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<tr>
	<tr>
		<td CLASS="LABEL">Network ID:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" size="10" MAXLENGTH="10" NAME="SearchCarrierNumber" VALUE=""></td>
		<td CLASS="LABEL" COLSPAN="2">Name:<br><input CLASS="LABEL" tabindex="2" TYPE="TEXT" NAME="SearchName" size="30" VALUE=""></td>
	</tr>
	</table>
</td>
<td VALIGN="TOP" rowspan="3">
	<table>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="12" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="13" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
	</table>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="9" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="10" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="11" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>


</form>
</body>
</html>
